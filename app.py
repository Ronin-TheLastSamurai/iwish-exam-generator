import os
import sys
import json
import re
import random
import base64
import datetime
import tempfile
import subprocess
from pathlib import Path

import streamlit as st
import pypandoc
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from google import genai
from google.genai import types

# ---------------------------------------------------------
# STREAMLIT PAGE CONFIGURATION
# ---------------------------------------------------------
st.set_page_config(
    page_title="iWish Exam Paper Generator",
    page_icon="🪄",
    layout="wide"
)

# ---------------------------------------------------------
# SECURE API KEY RESOLUTION
# ---------------------------------------------------------
def resolve_api_key() -> str:
    if "GEMINI_API_KEY" in st.secrets:
        return st.secrets["GEMINI_API_KEY"]
    if os.environ.get("GEMINI_API_KEY"):
        return os.environ.get("GEMINI_API_KEY")
    return ""

# ---------------------------------------------------------
# TOP 10 PRODUCTION MODEL FALLBACK HIERARCHY
# ---------------------------------------------------------
STATIC_MODELS_TO_TRY = [
    "gemini-3.8-flash",
    "gemini-3.7-flash",
    "gemini-3.6-flash",
    "gemini-3.1-pro-preview",
    "gemini-3.5-flash",
    "gemini-2.5-pro",
    "gemini-flash-latest",
    "gemini-pro-latest",
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
]

# ---------------------------------------------------------
# HINGLISH "MAGIC IN PROGRESS" MESSAGES
# ---------------------------------------------------------
TIER_1_MESSAGES = ["Mendeleev aur Newton ki aatma ko summon kiya jaa raha hai... 🧙‍♂️✨", "Chai ka cup utha lijiye, homework ka asli jaadu ab shuru ho raha hai... ☕🪄", "PDF ko AI ke hawaale kar diya hai... 🍿🚀"]
TIER_2_MESSAGES = ["Aapki PDF aise padhi jaa rahi hai jaise exam ki raat 3 baje padhai hoti hai... 📖⚡", "Topper ki tarah line-by-line scanning chalu hai... 🧐📄"]
TIER_3_MESSAGES = ["Aise homework questions dhoondh rahe hain jise dekh ke bachhe Google karna bhool jayein... 🤯📚", "Numerical aise set ho rahe hain jisme calculator bhi haath khade kar de... 🧮🔥"]
TIER_4_MESSAGES = ["Option B aur C mein thoda sa dimaag ka dahi karne wala twist daal rahe hain... 😈🎯", "Option elimination method ki dhajjiyan udane wala kaam chal raha hai... 🧠🌪️"]
TIER_5_MESSAGES = ["Homework ko ekdum masoom look de rahe hain taaki pehle lage ki kitna aasan hai... 😇📝", "Page layout aur fonts ekdum aesthetic set ho rahe hain... 🎨📄"]
TIER_6_MESSAGES = ["Boom! Jaadu complete. Ab mast evil smile ke saath homework bhej do! 🪄🎉", "Tadaaa! Masterpiece ban gaya, ab print nikaalo aur distribution shuru karo! 🎓🚀"]

# ---------------------------------------------------------
# CORE BACKEND FUNCTIONS & UNICODE TRANSLATION ENGINE
# ---------------------------------------------------------
def sanitize_raw_json_string(raw_json_str: str) -> str:
    cleaned = raw_json_str.strip()
    if cleaned.startswith("```json"): cleaned = cleaned[7:]
    elif cleaned.startswith("```"): cleaned = cleaned[3:]
    if cleaned.endswith("```"): cleaned = cleaned[:-3]
    return cleaned.strip()

def verify_assessment_json(data: dict, req_mcq: int, req_saq: int, req_laq: int) -> tuple[bool, str]:
    if not isinstance(data, dict): return False, "Response root must be a valid JSON object."
    
    mcqs = data.get("multiple_choice_questions", [])
    saqs = data.get("short_answer_questions", [])
    laqs = data.get("long_answer_questions", [])

    if len(mcqs) != req_mcq: return False, f"Expected exactly {req_mcq} MCQs."
    if len(saqs) != req_saq: return False, f"Expected exactly {req_saq} Short Answer questions."
    if len(laqs) != req_laq: return False, f"Expected exactly {req_laq} Long Answer questions."

    return True, "Valid"

def translate_to_unicode(text: str) -> str:
    """THE DETERMINISTIC UNICODE TRANSLATION ENGINE"""
    if not text: return ""

    # 1. Clean invisible characters and aggressively strip ALL dollar signs
    text = text.replace('\u00a0', ' ').replace('\u202f', ' ').replace('\u200b', '').replace('\ufeff', '')
    text = text.replace('$', '')
    
    # 2. Strip \text{} wrappers entirely and clean LaTeX artifacts
    text = re.sub(r'\\text\s*\{([^\}]+)\}', r'\1', text)
    text = re.sub(r'\\([()])', r'\1', text) # Unescape parentheses
    text = re.sub(r'\\?wedge\s*\{\s*-?1\s*\}', '⁻¹', text) # Remove wedge artifacts
    text = re.sub(r'\^?\\?wedge\s*\{\s*-?1\s*\}', '⁻¹', text)
    
    # 3. Standardize Chemical Arrows and Symbols
    text = text.replace(r'\rightarrow', '→').replace('-->', '→').replace('->', '→')
    text = text.replace(r'\rightleftharpoons', '⇌').replace('<=>', '⇌').replace('<->', '⇌')
    text = text.replace(r'\Delta', 'Δ').replace('Delta ', 'Δ ')
    text = text.replace(r'\circ', '°').replace('^o', '°')

    # 4. Standardize Enthalpy (Ruthlessly forces: ΔH = -XXX kJ mol⁻¹)
    text = re.sub(
        r'(?:Δ\s*)?H\s*=\s*(-?\d+(?:\.\d+)?)\s*(?:[kK]?[Jj](?:[\s~/\\]*mol\^?\{?-?1\}?|/mol)?\s*\}?\s*)+',
        r'ΔH = \1 kJ mol⁻¹',
        text
    )

    # 5. Translate Superscripts (Matches ^{3+} or ^3)
    def make_super(m):
        val = m.group(1).replace('o', '°').replace('\\circ', '°')
        return val.translate(str.maketrans("0123456789+-=()", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾"))
    text = re.sub(r'\^\{([^}]+)\}', make_super, text)
    text = re.sub(r'\^([0-9+-]+)', make_super, text)

    # 6. Translate Subscripts (Matches _{2} or _2)
    def make_sub(m):
        return m.group(1).translate(str.maketrans("0123456789+-=()aeox", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₒₓ"))
    text = re.sub(r'_\{([^}]+)\}', make_sub, text)
    text = re.sub(r'_([0-9a-zA-Z+-]+)', make_sub, text)

    # 7. Formatting fixes for Data Blocks (Molar Masses)
    text = re.sub(r'([A-Za-z₀-₉₍₎]+)\s*=\s*(\d)', r'\1 = \2', text) # Adds space around '='
    text = re.sub(r'(\d)\s*g\s*mol', r'\1 g mol', text)             # Adds space before 'g mol'
    text = re.sub(r',\s*([A-Z])', r', \1', text)                    # Adds space after comma

    # 8. Un-fuse specific English words if the AI squished them
    text = re.sub(r'(\d+(?:\.\d+)?)\s*molof\b', r'\1 mol of', text, flags=re.IGNORECASE)
    text = re.sub(r'(\d+(?:\.\d+)?)\s*mLof\b', r'\1 mL of', text, flags=re.IGNORECASE)
    text = re.sub(r'(\d+(?:\.\d+)?)\s*gof\b', r'\1 g of', text, flags=re.IGNORECASE)
    text = re.sub(r'\)and(\d)', r') and \1', text)

    # 9. THE BIDIRECTIONAL ISOLATOR: Forces equations onto their own lines safely
    text = re.sub(r'(equation:)\s+([A-Z0-9])', r'\1\n\n\2', text, flags=re.IGNORECASE)
    text = re.sub(r'(reaction:)\s+([A-Z0-9])', r'\1\n\n\2', text, flags=re.IGNORECASE)
    text = re.sub(r'(\([aqslg]\)|mol⁻¹|kJ)\s+([A-Z][a-z]+)', r'\1\n\n\2', text) # e.g., (g) In an experiment...
    
    text = re.sub(r' +', ' ', text)
    return text.strip()

# CACHED API POLLING
@st.cache_data(ttl=3600, show_spinner=False)
def get_available_models(api_key: str) -> list[str]:
    candidates = []
    try:
        client = genai.Client(api_key=api_key)
        for m in client.models.list():
            clean = m.name.replace("models/", "") if hasattr(m, "name") else ""
            if any(t in clean.lower() for t in ["flash", "pro"]) and not any(t in clean.lower() for t in ["vision", "tts", "clip", "transcribe", "image"]):
                candidates.append(clean)
    except Exception: pass

    combined = []
    for m in STATIC_MODELS_TO_TRY:
        if m not in combined: combined.append(m)
    for m in candidates:
        if m not in combined and "2.5-flash" not in m: combined.append(m)
    return combined

# ---------------------------------------------------------
# DYNAMIC PROMPT GENERATOR (STRICT UNICODE DIRECTIVE)
# ---------------------------------------------------------
def construct_australian_system_prompt(year_level: str, req_mcq: int, req_saq: int, req_laq: int) -> str:
    if year_level == "Year 6": curriculum_context = "FOCUS: States of matter, physical vs chemical changes, fair testing."
    elif year_level in ["Year 7", "Year 8"]: curriculum_context = "FOCUS: Particle theory, mixtures, separation techniques."
    elif year_level in ["Year 9", "Year 10"]: curriculum_context = "FOCUS: Atomic structure, balancing equations, reaction rates."
    else: curriculum_context = f"FOCUS: {year_level} ATAR Chemistry. Mole calculations, Equilibrium (Le Chatelier), Acids/Bases, Thermochemistry (SLC: 25°C, 100 kPa, V_m=24.79 L/mol)."

    return f"""You are a Chief Examination Item Writer for the Australian Curriculum.
Your objective is to craft an assessment for {year_level} based EXCLUSIVELY on [DOCUMENT 1: CHAPTER PDF]. Use [DOCUMENT 2: PYQ] only as a structure blueprint.

================================================================================
CRITICAL RULE: PLAIN TEXT UNICODE ONLY (NO LATEX MATH MODE)
================================================================================
- DO NOT USE DOLLAR SIGNS (`$`) EVER.
- DO NOT USE THE `\\text{{}}` COMMAND.
- ALWAYS put chemical equations on their own line! Hit Enter twice before and after any equation containing an arrow.
- Write chemical formulas strictly using underscores for subscripts and carets for superscripts.
  * CORRECT: H_2SO_4(aq)
  * CORRECT: Fe^{{3+}}
- Write reaction arrows as `->` or `<=>`.
  * CORRECT: 2H_2(g) + O_2(g) -> 2H_2O(l)
- Write enthalpy explicitly as: Delta H = -197 kJ mol^-1
- Do not write marks (e.g. "[2 marks]"). Do not generate dotted writing lines.

================================================================================
REQUIRED JSON SCHEMA
================================================================================
Output STRICT JSON exactly matching:
{{
  "multiple_choice_questions": [
    {{"number": 1, "stem": "Text", "options": {{"A": "Text", "B": "Text", "C": "Text", "D": "Text"}}}}
  ],
  "short_answer_questions": [
    {{"number": 1, "stem": "Text", "sub_parts": ["a) Text", "b) Text"]}}
  ],
  "long_answer_questions": [
    {{"number": 1, "stem": "Text", "sub_parts": ["a) Text", "b) Text", "c) Text"]}}
  ]
}}
GENERATE EXACTLY: {req_mcq} MCQs, {req_saq} SAQs, {req_laq} LAQs."""

def generate_and_verify_homework(chapter_path: str, pyq_path: str | None, api_key: str, year_level: str, difficulty: str, custom_keywords: str, req_mcq: int, req_saq: int, req_laq: int) -> dict:
    client = genai.Client(api_key=api_key)
    
    uploaded_chapter = client.files.upload(file=chapter_path)
    uploaded_pyq = client.files.upload(file=pyq_path) if pyq_path and os.path.exists(pyq_path) else None

    system_prompt = construct_australian_system_prompt(year_level, req_mcq, req_saq, req_laq)
    user_prompt = f"Generate {year_level} paper.\nDIFFICULTY: {difficulty.upper()}\nKEYWORDS: {custom_keywords}\nEnsure NO DOLLAR SIGNS. Output valid JSON only."

    try:
        model_errors = {}
        for model_name in get_available_models(api_key):
            contents = [uploaded_chapter]
            if uploaded_pyq: contents.append(uploaded_pyq)
            contents.append(user_prompt)

            for _ in range(2):
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=contents,
                        config=types.GenerateContentConfig(system_instruction=system_prompt, temperature=0.2, response_mime_type="application/json")
                    )
                    parsed_json = json.loads(sanitize_raw_json_string(response.text))
                    is_valid, msg = verify_assessment_json(parsed_json, req_mcq, req_saq, req_laq)
                    if is_valid: return parsed_json
                    contents.extend([response.text, f"VALIDATION ERROR: {msg}\nRegenerate."])
                except Exception as err:
                    model_errors[model_name] = str(err)
                    break

        raise RuntimeError(f"All models failed:\n" + "\n".join([f"• {m}: {e}" for m, e in model_errors.items()]))
    finally:
        try: client.files.delete(name=uploaded_chapter.name)
        except: pass
        if uploaded_pyq:
            try: client.files.delete(name=uploaded_pyq.name)
            except: pass

def json_to_markdown(data: dict, selected_date: str) -> str:
    md_lines = ["# Homework\n", f"(Date - {selected_date})\n"]
    for q_type, title in [("multiple_choice_questions", "MULTIPLE CHOICE QUESTIONS"), ("short_answer_questions", "SHORT ANSWER QUESTIONS"), ("long_answer_questions", "LONG ANSWER QUESTIONS")]:
        questions = data.get(q_type, [])
        if questions:
            md_lines.append(f"## {title}\n")
            for q in questions:
                num = q.get("number", 1)
                stem = translate_to_unicode(str(q.get("stem", "")).strip())
                md_lines.append(f"**{num}.** {stem}\n")
                if q_type == "multiple_choice_questions":
                    for opt_key in ["A", "B", "C", "D"]:
                        opt_val = re.sub(rf'^{opt_key}[\)\.\:\-\s]+\s*', '', str(q.get("options", {}).get(opt_key, "")).strip(), flags=re.IGNORECASE)
                        md_lines.append(f"**{opt_key})** {translate_to_unicode(opt_val)}\n")
                else:
                    for sub in q.get("sub_parts", []):
                        sub_formatted = re.sub(r'^([a-d]\)|[ivx]+\))\s*', r'**\1** ', translate_to_unicode(str(sub).strip()), flags=re.IGNORECASE)
                        md_lines.append(f"{sub_formatted if sub_formatted.startswith('**') else '**-** ' + sub_formatted}\n")
                md_lines.append("")
    return "\n".join(md_lines).strip()

def build_base_docx(markdown_text: str, output_docx_path: str):
    # CRITICAL: "--from=markdown" completely disables Pandoc's LaTeX math compiler!
    pypandoc.convert_text(source=markdown_text, to="docx", format="markdown", outputfile=output_docx_path, extra_args=["--from=markdown", "--to=docx"])

def style_assessment_docx(docx_path: str, logo_img_path: str = "iwish_logo.jpg"):
    doc = Document(docx_path)
    section = doc.sections[0]
    section.top_margin = Inches(0.8); section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.9); section.right_margin = Inches(0.9)
    
    header_p = section.header.paragraphs[0]
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(logo_img_path): header_p.add_run().add_picture(logo_img_path, width=Inches(1.8))

    footer_p = section.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r1 = footer_p.add_run("Page ")
    r1.font.name = "Arial"; r1.font.size = Pt(8.5); r1.font.color.rgb = RGBColor(100, 116, 139)
    footer_p._p.append(parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w')))
    r2 = footer_p.add_run(" of ")
    r2.font.name = "Arial"; r2.font.size = Pt(8.5); r2.font.color.rgb = RGBColor(100, 116, 139)
    footer_p._p.append(parse_xml(r'<w:fldSimple %s w:instr="NUMPAGES"/>' % nsdecls('w')))

    section_keywords = ["MULTIPLE CHOICE QUESTIONS", "SHORT ANSWER QUESTIONS", "LONG ANSWER QUESTIONS"]
    is_first_section = True

    for idx, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text: continue
        clean_upper = text.upper().replace("#", "").strip()

        if text.startswith("Homework") and len(text) < 15:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(2); p.paragraph_format.keep_with_next = True
            for r in p.runs: r.font.name = "Arial"; r.font.size = Pt(14); r.font.bold = True
            continue

        if text.startswith("(Date -"):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(14); p.paragraph_format.keep_with_next = True
            for r in p.runs: r.font.name = "Arial"; r.font.size = Pt(10); r.font.color.rgb = RGBColor(71, 85, 105)
            continue

        matched_section = next((s for s in section_keywords if s in clean_upper), None)
        if matched_section:
            p.text = matched_section
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in p.runs: run.font.name = "Arial"; run.font.size = Pt(13.5); run.font.bold = True; run.font.underline = True; run.font.color.rgb = RGBColor(30, 41, 59)
            if not is_first_section:
                p._p.addprevious(parse_xml(r'<w:p %s><w:pPr><w:pBrd><w:bottom w:val="single" w:sz="12" w:space="1" w:color="CBD5E1"/></w:pBrd><w:spacing w:before="240" w:after="160"/></w:pPr><w:r><w:rPr><w:sz w:val="4"/></w:rPr><w:t> </w:t></w:r></w:p>' % nsdecls('w')))
                p.paragraph_format.space_before = Pt(6)
            else: p.paragraph_format.space_before = Pt(8); is_first_section = False
            p.paragraph_format.space_after = Pt(8); p.paragraph_format.keep_with_next = True; p.paragraph_format.keep_together = True
            continue

        # SMART CENTERING: Only center if it contains an arrow AND is isolated (short length, no plain English prose)
        if ('→' in text or '⇌' in text) and len(text) < 150 and "according to" not in text and "experiment" not in text:
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(6)
            continue

        mcq_opt_match = re.match(r'^([A-D]\))\s*(.*)', text)
        if mcq_opt_match:
            p.paragraph_format.left_indent = Inches(0.18); p.paragraph_format.keep_together = True
            p.paragraph_format.space_before = Pt(1); p.paragraph_format.space_after = Pt(12) if mcq_opt_match.group(1).upper().startswith('D') else Pt(2)
            continue

        sub_part_match = re.match(r'^([a-d]\)|[ivx]+\))\s*(.*)', text, re.IGNORECASE)
        if sub_part_match:
            p.paragraph_format.left_indent = Inches(0.12); p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(4)
            continue

        if re.match(r'^(\d+[\.\)])\s*(.*)', text):
            p.paragraph_format.left_indent = Inches(0); p.paragraph_format.space_before = Pt(10); p.paragraph_format.space_after = Pt(3)
            continue

    doc.save(docx_path)

def convert_docx_to_pdf(docx_path: str, output_pdf_path: str) -> bool:
    out_dir = os.path.dirname(output_pdf_path)
    try:
        subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", out_dir], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True, timeout=60)
        if os.path.exists(os.path.join(out_dir, Path(docx_path).stem + ".pdf")):
            os.replace(os.path.join(out_dir, Path(docx_path).stem + ".pdf"), output_pdf_path)
            return True
    except: pass
    return False

# ---------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------
if "assessment_data" not in st.session_state: st.session_state["assessment_data"] = None

st.title("✨ iWish Exam Paper Generator ✨")
st.markdown("Generate authentic Curriculum & ATAR homework assessments from textbook PDFs.")

api_key = resolve_api_key() or st.sidebar.text_input("Enter Gemini API Key", type="password")
year_level = st.sidebar.selectbox("🎓 Grade / Year Level", ["Year 6", "Year 7", "Year 8", "Year 9", "Year 10", "Year 11 (ATAR)", "Year 12 (ATAR)"], index=5)
selected_date = st.sidebar.date_input("Homework Date", value=datetime.date.today()).strftime("%d.%m.%Y")
difficulty = st.sidebar.select_slider("📊 Difficulty Level", ["Easy", "Medium", "Difficult"], value="Medium")

req_mcq = st.sidebar.slider("Multiple Choice Questions", 0, 10, 5)
req_saq = st.sidebar.slider("Short Answer Questions", 0, 10, 5)
req_laq = st.sidebar.slider("Long Answer Questions", 0, 10, 0)
custom_keywords = st.sidebar.text_area("🎯 Target Specific Subtopics (Optional):") if st.sidebar.checkbox("Target Specific Subtopics / Keywords", False) else ""

col_upload1, col_upload2 = st.columns(2)
with col_upload1: uploaded_chapter = st.file_uploader("📘 Upload Chapter PDF (Mandatory)", type=["pdf"])
with col_upload2: uploaded_pyq = st.file_uploader("📝 Upload PYQ (Optional Style Blueprint)", type=["pdf"])

col_action1, col_action2 = st.columns([3, 1])
with col_action1: generate_clicked = st.button("🪄 Generate Assessment Paper", type="primary", use_container_width=True)
with col_action2:
    if st.session_state["assessment_data"] and st.button("🔄 Clear Exam", use_container_width=True):
        st.session_state["assessment_data"] = None
        st.rerun()

if generate_clicked:
    if not api_key: st.error("API Key required."); st.stop()
    if req_mcq + req_saq + req_laq == 0: st.error("Select at least 1 question type."); st.stop()
    if not uploaded_chapter: st.error("Chapter PDF required."); st.stop()

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_chapter_pdf = os.path.join(temp_dir, uploaded_chapter.name)
        with open(temp_chapter_pdf, "wb") as f: f.write(uploaded_chapter.getbuffer())
        temp_pyq_pdf = os.path.join(temp_dir, uploaded_pyq.name) if uploaded_pyq else None
        if temp_pyq_pdf: 
            with open(temp_pyq_pdf, "wb") as f: f.write(uploaded_pyq.getbuffer())

        base_stem = Path(uploaded_chapter.name).stem
        docx_output_path = os.path.join(temp_dir, f"{base_stem}_Assessment.docx")
        pdf_output_path = os.path.join(temp_dir, f"{base_stem}_Assessment.pdf")

        progress_bar = st.progress(0); status_box = st.empty()

        try:
            progress_bar.progress(20); status_box.info(random.choice(TIER_1_MESSAGES))
            assessment_json = generate_and_verify_homework(temp_chapter_pdf, temp_pyq_pdf, api_key, year_level, difficulty, custom_keywords, req_mcq, req_saq, req_laq)
            
            progress_bar.progress(50); status_box.info(random.choice(TIER_3_MESSAGES))
            markdown_content = json_to_markdown(assessment_json, selected_date)
            
            progress_bar.progress(70); status_box.info(random.choice(TIER_5_MESSAGES))
            build_base_docx(markdown_content, docx_output_path)
            style_assessment_docx(docx_output_path)
            pdf_success = convert_docx_to_pdf(docx_output_path, pdf_output_path)

            progress_bar.progress(100); status_box.success(random.choice(TIER_6_MESSAGES))

            with open(docx_output_path, "rb") as f_docx: docx_bytes = f_docx.read()
            pdf_bytes = open(pdf_output_path, "rb").read() if pdf_success and os.path.exists(pdf_output_path) else None

            st.session_state["assessment_data"] = {"docx_bytes": docx_bytes, "pdf_bytes": pdf_bytes, "base_stem": base_stem}
            status_box.empty(); progress_bar.empty()
            st.rerun()

        except Exception as e:
            status_box.empty(); progress_bar.empty()
            st.error(f"Generation error: {e}")

if st.session_state["assessment_data"]:
    data = st.session_state["assessment_data"]
    st.success("🎉 Assessment complete! Download below.")
    col1, col2 = st.columns(2)
    with col1: st.download_button("📥 Download DOCX", data["docx_bytes"], f"{data['base_stem']}_Assessment.docx", use_container_width=True)
    with col2:
        if data["pdf_bytes"]: st.download_button("📥 Download PDF", data["pdf_bytes"], f"{data['base_stem']}_Assessment.pdf", use_container_width=True)
        else: st.info("PDF engine unavailable.")
    if data["pdf_bytes"]:
        st.markdown(f'<iframe src="data:application/pdf;base64,{base64.b64encode(data["pdf_bytes"]).decode("utf-8")}" width="100%" height="850" type="application/pdf"></iframe>', unsafe_allow_html=True)
