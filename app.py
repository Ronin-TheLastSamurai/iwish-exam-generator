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
from google.genai.errors import APIError

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
TIER_1_MESSAGES = [
    "Mendeleev aur Newton ki aatma ko summon kiya jaa raha hai... 🧙‍♂️✨",
    "Chai ka cup utha lijiye, homework ka asli jaadu ab shuru ho raha hai... ☕🪄",
    "PDF ko AI ke hawaale kar diya hai, ab bas aage ka tamasha dekhiye... 🍿🚀",
    "Teacher mode: ON! Bachhon ki azaadi bas kuch der ki mehmaan hai... 😈⏰"
]

TIER_2_MESSAGES = [
    "Aapki PDF aise padhi jaa rahi hai jaise exam ki raat 3 baje padhai hoti hai... 📖⚡",
    "AI poori PDF ko ghol ke pee raha hai, ek-ek concept dhoondh raha hai... 🧪🔍",
    "Topper ki tarah line-by-line scanning chalu hai taaki kuch na chhoote... 🧐📄"
]

TIER_3_MESSAGES = [
    "Aise homework questions dhoondh rahe hain jise dekh ke bachhe Google karna bhool jayein... 🤯📚",
    "ChatGPT bhi do baar soche aur 'I cannot answer' keh de, aise sawaal ban rahe hain... 🤖💥",
    "Numerical aise set ho rahe hain jisme calculator bhi haath khade kar de... 🧮🔥"
]

TIER_4_MESSAGES = [
    "Option B aur C mein thoda sa dimaag ka dahi karne wala twist daal rahe hain... 😈🎯",
    "Saare options ek jaise dikhein, aisi subtle ninja technique lagayi jaa rahi hai... 🥷✨",
    "Option elimination method ki dhajjiyan udane wala kaam chal raha hai... 🧠🌪️"
]

TIER_5_MESSAGES = [
    "Homework ko ekdum masoom look de rahe hain taaki pehle lage ki kitna aasan hai... 😇📝",
    "Page layout, clean fonts aur iWish logo ekdum aesthetic set ho rahe hain... 🎨📄",
    "Equations aur chemical formulas ko sundar sa LaTeX makeup lagaya jaa raha hai... 💄⚗️"
]

TIER_6_MESSAGES = [
    "Boom! Jaadu complete. Ab mast evil smile ke saath homework bhej do! 🪄🎉",
    "Paper ready hai! Ab WhatsApp group pe bhej ke phone silent kar lijiye... 📱🔥",
    "Tadaaa! Masterpiece ban gaya, ab print nikaalo aur distribution shuru karo! 🎓🚀"
]

# ---------------------------------------------------------
# CORE BACKEND FUNCTIONS & LIGHTNING-FAST PRE-PROCESSOR
# ---------------------------------------------------------
def sanitize_raw_json_string(raw_json_str: str) -> str:
    cleaned = raw_json_str.strip()
    if cleaned.startswith("```json"):
        cleaned = cleaned[7:]
    elif cleaned.startswith("```"):
        cleaned = cleaned[3:]
    if cleaned.endswith("```"):
        cleaned = cleaned[:-3]
    cleaned = cleaned.strip()

    latex_cmd_regex = (
        r'(?<!\\)\\(?=(?:text|textbf|textrm|textit|tfrac|frac|times|tau|theta|to|'
        r'rightarrow|rightleftharpoons|right|rho|rangle|rm|nu|nabla|neq|not|neg|'
        r'beta|bar|bf|bm|begin|end|boldsymbol|bullet|Delta|alpha|gamma|sigma|'
        r'lambda|mu|pi|phi|omega|degree|circ|pm|mp|le|ge|approx|sim|equiv|'
        r'cdot|dots|cdots|left|sqrt|partial|sum|int|infty)\b)'
    )
    return re.sub(latex_cmd_regex, r'\\\\', cleaned)


def verify_assessment_json(data: dict, req_mcq: int, req_saq: int, req_laq: int) -> tuple[bool, str]:
    if not isinstance(data, dict):
        return False, "Response root must be a valid JSON object."

    mcqs = data.get("multiple_choice_questions", [])
    saqs = data.get("short_answer_questions", [])
    laqs = data.get("long_answer_questions", [])

    if len(mcqs) != req_mcq: return False, f"Expected exactly {req_mcq} MCQs."
    if len(saqs) != req_saq: return False, f"Expected exactly {req_saq} Short Answer questions."
    if len(laqs) != req_laq: return False, f"Expected exactly {req_laq} Long Answer questions."

    for i, q in enumerate(mcqs, 1):
        if not str(q.get("stem", "")).strip(): return False, f"MCQ #{i} has an empty stem."
        opts = q.get("options", {})
        if not isinstance(opts, dict): return False, f"MCQ #{i} options must be a dictionary."
        for opt_key in ["A", "B", "C", "D"]:
            if opt_key not in opts or not str(opts[opt_key]).strip(): return False, f"MCQ #{i} is missing option '{opt_key}'."

    for i, q in enumerate(saqs, 1):
        if not str(q.get("stem", "")).strip(): return False, f"Short Answer #{i} has an empty stem."
        for sub in q.get("sub_parts", []):
            if not str(sub).strip(): return False, f"Short Answer #{i} has an empty sub-part."

    for i, q in enumerate(laqs, 1):
        if not str(q.get("stem", "")).strip(): return False, f"Long Answer #{i} has an empty stem."
        sub_parts = q.get("sub_parts", [])
        if not sub_parts or not isinstance(sub_parts, list): return False, f"Long Answer #{i} must include sub_parts."

    return True, "Valid"


def repair_math_syntax(text: str) -> str:
    """Nuclear Pre-processor: Extremely fast regex logic to repair LaTeX before Pandoc compiles."""
    if not text: return ""

    # 1. Normalize unicode and un-escape parentheses
    text = text.replace('\u00a0', ' ').replace('\u202f', ' ').replace('\u200b', '').replace('\ufeff', '')
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    text = re.sub(r'\\([()])', r'\1', text)

    # 2. GREEDY ENTHALPY FIX: Ruthlessly swallow duplicate/mangled units up to the next valid word
    text = re.sub(r'\^?\{\s*-?1\s*\}\s*\\wedge\s*\{\s*-?1\s*\}', '^{-1}', text)
    text = re.sub(
        r'\$?\s*(?:with\s+)?(?:\\?[Dd]elta\s*)?H\s*=\s*(-?\d+(?:\.\d+)?)\s*(?:\\text\s*\{)?\s*(?:[kK]?[Jj](?:[\s~]*mol\^?\{?-?1\}?|/mol)?\s*\}?\s*)+\$?',
        r' $\\Delta H = \1\\text{ kJ mol}^{-1}$ ',
        text
    )

    # 3. Empty Base Box Issue (L$^{-1}$ or L^{-1} -> $\text{L}^{-1}$)
    text = re.sub(r'\b([A-Za-z]+)\s*\$\^', r'$\1^', text)
    text = re.sub(r'\b([A-Za-z]+)\s*\^\{?(-?\d+)\}?', r'$\1^{\2}$', text)

    if text.count('$') % 2 != 0: text += '$'

    # 4. FAST Force-wrapping of full reaction equations (Line-by-line, zero backtracking)
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if any(arr in line for arr in [r'\rightarrow', r'\rightleftharpoons', r'\to', '→', '⇌']):
            stripped = line.strip()
            if not (stripped.startswith('$$') and stripped.endswith('$$')):
                clean_line = line.replace('$', '') # Avoid nested dollars
                lines[i] = f"$${clean_line.strip()}$$"
    text = '\n'.join(lines)

    # 5. Process Outside vs Inside Math Mode
    text = text.replace('$$', '___DOUBLEDOLLAR___')
    parts = text.split('$')

    for i in range(len(parts)):
        if i % 2 == 0:
            # OUTSIDE MATH MODE: Fast replacement for orphaned macros
            chunk = parts[i]
            # Wrap \text{...}_... into math mode
            chunk = re.sub(r'(\\[a-zA-Z]+(?:\{[^\}]*\})?_\{?[A-Za-z0-9+-]*\}?(?:\([aqslg]+\))?)', r'$\1$', chunk)
            # Catch orphaned chemical formulas with underscores (e.g. AgNO_3(aq)) to prevent markdown italic swallowing
            chunk = re.sub(r'(?<![\$\\a-zA-Z])([A-Z][A-Za-z0-9\(\)\[\]]*_\{?[a-zA-Z0-9+-]+\}?(?:\([aqslg]+\))?)(?![\$a-zA-Z])', r'$\1$', chunk)
            parts[i] = chunk
        else:
            # INSIDE MATH MODE: Un-fuse words by actively injecting \text{ } around English vocabulary
            chunk = parts[i]
            
            # Protect existing \text{...} to prevent nesting errors (\text{ kJ \text{ mol } })
            protected_texts = []
            def stash_text(m):
                protected_texts.append(m.group(0))
                return f"___TEXTBLOCK_{len(protected_texts)-1}___"
            
            chunk = re.sub(r'\\text\s*\{[^\}]+\}', stash_text, chunk)
            
            # Apply word unfuser safely
            stopwords = r'\b(mol|mL|g|L|kg|kPa|atm|of|and|is|with|mixed|excess|solution|precipitate|mass|yield|sample|solid|formed|produced|reacts|the|in|to|from|contains|gas|at|conditions|determine|calculate|which|reactant|volume)\b'
            chunk = re.sub(stopwords, r'\\text{ \1 }', chunk, flags=re.IGNORECASE)
            
            # Restore \text{...}
            for j, block in enumerate(protected_texts):
                chunk = chunk.replace(f"___TEXTBLOCK_{j}___", block)
            
            parts[i] = chunk

    text = '$'.join(parts)
    text = text.replace('___DOUBLEDOLLAR___', '$$')

    # 6. Final spacing cleanup
    text = re.sub(r'\$\s*\$', '', text)
    text = re.sub(r' +', ' ', text)

    if text.count('$') % 2 != 0: text += '$'
    return text.strip()


# CACHED API POLLING: Only runs once per hour, stopping the pre-generation delay
@st.cache_data(ttl=3600, show_spinner=False)
def get_available_models(api_key: str) -> list[str]:
    candidates = []
    try:
        client = genai.Client(api_key=api_key)
        models_pager = client.models.list()
        for m in models_pager:
            clean_name = m.name.replace("models/", "") if hasattr(m, "name") else ""
            if any(tag in clean_name.lower() for tag in ["flash", "pro"]) and not any(tag in clean_name.lower() for tag in ["vision", "tts", "clip", "transcribe", "image"]):
                candidates.append(clean_name)
    except Exception:
        pass

    combined = []
    for m in STATIC_MODELS_TO_TRY:
        if m not in combined: combined.append(m)
    for m in candidates:
        if m not in combined and "2.5-flash" not in m: combined.append(m)

    return combined


# ---------------------------------------------------------
# DYNAMIC AUSTRALIAN CURRICULUM PROMPT GENERATOR
# ---------------------------------------------------------
def construct_australian_system_prompt(year_level: str, req_mcq: int, req_saq: int, req_laq: int) -> str:
    if year_level == "Year 6":
        curriculum_context = """CURRICULUM CONTEXT: Australian Curriculum (ACARA v9.0) - Year 6 Primary Science.
- FOCUS: Observable properties, states of matter (solid, liquid, gas), reversible physical changes vs. irreversible chemical changes.
- SCIENCE INQUIRY SKILLS: Fair testing principles. Reading simple data tables."""
    elif year_level in ["Year 7", "Year 8"]:
        curriculum_context = """CURRICULUM CONTEXT: Australian Curriculum (ACARA v9.0) - Years 7-8 Junior Secondary Science.
- FOCUS: Particle model of matter, Pure substances vs. mixtures, Separation techniques, Physical vs. chemical changes.
- SCIENCE INQUIRY SKILLS: Designing fair tests, writing hypotheses, identifying laboratory errors."""
    elif year_level in ["Year 9", "Year 10"]:
        curriculum_context = """CURRICULUM CONTEXT: Australian Curriculum (ACARA v9.0) - Years 9-10 Middle Secondary Science.
- FOCUS: Atomic structure, Periodic Table organization, Chemical reactions & conservation of mass, Rates of reaction.
- SCIENCE INQUIRY SKILLS: Distinguishing between accuracy, reliability, and validity."""
    else:  # Year 11 (ATAR) or Year 12 (ATAR)
        curriculum_context = f"""CURRICULUM CONTEXT: Australian Senior Secondary Chemistry ({year_level} ATAR).
- FOCUS: Quantitative chemistry, Gas stoichiometry (SLC: 25°C / 298.15 K and 100 kPa, V_m = 24.79 L/mol), Equilibrium, Acids and bases, Thermochemistry, Organic chemistry.
- AUSTRALIAN ATAR COMMAND VERBS: Explain, Justify, Assess, Evaluate, Deduce.
- WORKING SCIENTIFICALLY: Experimental rigor, titration error analysis, glassware rinsing procedures."""

    return f"""You are a Chief Examination Item Writer and Senior Curriculum Assessor for the Australian Secondary School System.
Your objective is to craft an authentic, curriculum-aligned Australian Examination Paper for {year_level} based EXCLUSIVELY on the provided Chapter PDF.

================================================================================
1. STRICT CONTENT BOUNDARY
================================================================================
- [DOCUMENT 1: CHAPTER PDF] defines the absolute syllabus boundary.
- If [DOCUMENT 2: PYQ] is provided, use it SOLELY as a structural blueprint. Do NOT copy past paper questions verbatim.

================================================================================
2. CURRICULUM STANDARDS FOR {year_level.upper()}
================================================================================
{curriculum_context}

================================================================================
3. STRICT FORMATTING CONSTRAINTS
================================================================================
- NO MARKS: Never write "[2 marks]", "(3 marks)".
- NO LINES: Do not generate response lines or writing spaces.
- NO META-LANGUAGE: Never write "refer to the text".
- WRAP ALL CHEMICAL REACTIONS IN DOUBLE DOLLAR SIGNS: `$$ \\text{{AgNO}}_3(aq) + \\text{{NaCl}}(aq) \\rightarrow ... $$`

================================================================================
4. REQUIRED JSON SCHEMA
================================================================================
Output STRICT, raw, parsable JSON matching this exact structure:
{{
  "multiple_choice_questions": [
    {{
      "number": 1,
      "stem": "Question stem",
      "options": {{"A": "Option", "B": "Option", "C": "Option", "D": "Option"}}
    }}
  ],
  "short_answer_questions": [
    {{
      "number": 1,
      "stem": "Scenario",
      "sub_parts": ["a) Prompt", "b) Prompt"]
    }}
  ],
  "long_answer_questions": [
    {{
      "number": 1,
      "stem": "Investigation",
      "sub_parts": ["a) Prompt", "b) Prompt", "c) Prompt"]
    }}
  ]
}}
EXACT COUNTS TO GENERATE: {req_mcq} MCQs, {req_saq} SAQs, {req_laq} LAQs."""


def generate_and_verify_homework(chapter_pdf_path: str, pyq_pdf_path: str | None, api_key: str, year_level: str, difficulty: str, custom_keywords: str, req_mcq: int, req_saq: int, req_laq: int) -> dict:
    client = genai.Client(api_key=api_key)
    
    uploaded_chapter = client.files.upload(file=chapter_pdf_path)
    uploaded_pyq = client.files.upload(file=pyq_pdf_path) if pyq_pdf_path and os.path.exists(pyq_pdf_path) else None

    system_prompt = construct_australian_system_prompt(year_level, req_mcq, req_saq, req_laq)
    difficulty_instruction = f"DIFFICULTY LEVEL: {difficulty.upper()}.\n"
    keywords_instruction = f"FOCUS KEYWORDS: {custom_keywords}\n" if custom_keywords.strip() else ""

    user_prompt = f"""Generate the {year_level} assessment paper in strict accordance with the Australian Curriculum framework.
{difficulty_instruction}
{keywords_instruction}
Output valid JSON only."""

    try:
        model_errors = {}
        active_models = get_available_models(api_key)

        for model_name in active_models:
            conversation_contents = [uploaded_chapter]
            if uploaded_pyq: conversation_contents.append(uploaded_pyq)
            conversation_contents.append(user_prompt)

            for _ in range(2):
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=conversation_contents,
                        config=types.GenerateContentConfig(
                            system_instruction=system_prompt,
                            temperature=0.2,
                            response_mime_type="application/json"
                        )
                    )

                    sanitized = sanitize_raw_json_string(response.text)
                    parsed_json = json.loads(sanitized)
                    is_valid, validation_msg = verify_assessment_json(parsed_json, req_mcq, req_saq, req_laq)

                    if is_valid: return parsed_json

                    conversation_contents.extend([response.text, f"VALIDATION ERROR: {validation_msg}\nRegenerate."])
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
    for q_type, title in [("multiple_choice_questions", "MULTIPLE CHOICE QUESTIONS"), 
                          ("short_answer_questions", "SHORT ANSWER QUESTIONS"), 
                          ("long_answer_questions", "LONG ANSWER QUESTIONS")]:
        questions = data.get(q_type, [])
        if questions:
            md_lines.append(f"## {title}\n")
            for q in questions:
                num = q.get("number", 1)
                stem = repair_math_syntax(str(q.get("stem", "")).strip())
                md_lines.append(f"**{num}.** {stem}\n")
                if q_type == "multiple_choice_questions":
                    for opt_key in ["A", "B", "C", "D"]:
                        opt_val = re.sub(rf'^{opt_key}[\)\.\:\-\s]+\s*', '', str(q.get("options", {}).get(opt_key, "")).strip(), flags=re.IGNORECASE)
                        md_lines.append(f"**{opt_key})** {repair_math_syntax(opt_val)}\n")
                else:
                    for sub in q.get("sub_parts", []):
                        sub_formatted = re.sub(r'^([a-d]\)|[ivx]+\))\s*', r'**\1** ', repair_math_syntax(str(sub).strip()), flags=re.IGNORECASE)
                        md_lines.append(f"{sub_formatted if sub_formatted.startswith('**') else '**-** ' + sub_formatted}\n")
                md_lines.append("")
    return "\n".join(md_lines).strip()


def build_base_docx(markdown_text: str, output_docx_path: str):
    pypandoc.convert_text(source=markdown_text, to="docx", format="markdown", outputfile=output_docx_path, extra_args=["--from=markdown+tex_math_dollars+tex_math_single_backslash", "--to=docx"])


def style_assessment_docx(docx_path: str, logo_img_path: str = "iwish_logo.jpg"):
    doc = Document(docx_path)
    section = doc.sections[0]
    section.top_margin = Inches(0.8); section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.9); section.right_margin = Inches(0.9)
    section.header_distance = Inches(0.35); section.footer_distance = Inches(0.4)

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
    valid_paragraphs = [p for p in doc.paragraphs if p.text.strip() or p._p.xpath('.//m:oMathPara') or p._p.xpath('.//m:oMath')]

    for idx, p in enumerate(valid_paragraphs):
        text = p.text.strip()
        clean_upper = text.upper().replace("#", "").strip()
        next_p = valid_paragraphs[idx + 1] if idx + 1 < len(valid_paragraphs) else None
        next_text = next_p.text.strip() if next_p else ""
        is_terminal = next_p is None or any(s in next_text.upper() for s in section_keywords) or bool(re.match(r'^\d+[\.\)]\s*', next_text))

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
            else:
                p.paragraph_format.space_before = Pt(8); is_first_section = False
            p.paragraph_format.space_after = Pt(8); p.paragraph_format.keep_with_next = True; p.paragraph_format.keep_together = True
            continue

        mcq_opt_match = re.match(r'^([A-D]\))\s*(.*)', text)
        if mcq_opt_match:
            p.paragraph_format.left_indent = Inches(0.18); p.paragraph_format.keep_together = True
            p.paragraph_format.space_before = Pt(1); p.paragraph_format.space_after = Pt(12) if mcq_opt_match.group(1).upper().startswith('D') or is_terminal else Pt(2)
            p.paragraph_format.keep_with_next = not (mcq_opt_match.group(1).upper().startswith('D') or is_terminal)
            continue

        sub_part_match = re.match(r'^([a-d]\)|[ivx]+\))\s*(.*)', text, re.IGNORECASE)
        if sub_part_match:
            p.paragraph_format.left_indent = Inches(0.12); p.paragraph_format.space_before = Pt(2); p.paragraph_format.keep_together = True
            p.paragraph_format.space_after = Pt(12) if is_terminal else Pt(4)
            p.paragraph_format.keep_with_next = not is_terminal
            continue

        if re.match(r'^(\d+[\.\)])\s*(.*)', text):
            p.paragraph_format.left_indent = Inches(0); p.paragraph_format.space_before = Pt(10); p.paragraph_format.keep_together = True
            p.paragraph_format.space_after = Pt(12) if is_terminal else Pt(3)
            p.paragraph_format.keep_with_next = not is_terminal
            continue

        if bool(p._p.xpath('.//m:oMathPara') or (p._p.xpath('.//m:oMath') and len(text) < 5)):
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(6); p.paragraph_format.keep_together = True
            p.paragraph_format.space_after = Pt(12) if is_terminal else Pt(6)
            p.paragraph_format.keep_with_next = not is_terminal
            continue

        p.paragraph_format.space_before = Pt(2); p.paragraph_format.keep_together = True
        p.paragraph_format.space_after = Pt(12) if is_terminal else Pt(4)
        p.paragraph_format.keep_with_next = not is_terminal

    doc.save(docx_path)


def convert_docx_to_pdf(docx_path: str, output_pdf_path: str) -> bool:
    out_dir = os.path.dirname(output_pdf_path)
    try:
        subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", out_dir], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True, timeout=60)
        expected_pdf = os.path.join(out_dir, Path(docx_path).stem + ".pdf")
        if os.path.exists(expected_pdf):
            if expected_pdf != output_pdf_path: os.replace(expected_pdf, output_pdf_path)
            return True
    except: pass
    return False

# ---------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------
if "assessment_data" not in st.session_state: st.session_state["assessment_data"] = None

st.title("✨ iWish Exam Paper Generator ✨")
st.markdown("Generate authentic Australian Curriculum & ATAR homework assessments from textbook PDFs.")

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
        temp_pyq_pdf = None
        if uploaded_pyq:
            temp_pyq_pdf = os.path.join(temp_dir, uploaded_pyq.name)
            with open(temp_pyq_pdf, "wb") as f: f.write(uploaded_pyq.getbuffer())

        base_stem = Path(uploaded_chapter.name).stem
        docx_output_path = os.path.join(temp_dir, f"{base_stem}_Assessment.docx")
        pdf_output_path = os.path.join(temp_dir, f"{base_stem}_Assessment.pdf")

        progress_bar = st.progress(0); status_box = st.empty()

        try:
            progress_bar.progress(10); status_box.info(random.choice(TIER_1_MESSAGES))
            progress_bar.progress(25); status_box.info(random.choice(TIER_2_MESSAGES))

            assessment_json = generate_and_verify_homework(temp_chapter_pdf, temp_pyq_pdf, api_key, year_level, difficulty, custom_keywords, req_mcq, req_saq, req_laq)

            progress_bar.progress(50); status_box.info(random.choice(TIER_3_MESSAGES))
            markdown_content = json_to_markdown(assessment_json, selected_date)
            
            progress_bar.progress(68); status_box.info(random.choice(TIER_4_MESSAGES))
            build_base_docx(markdown_content, docx_output_path)
            
            progress_bar.progress(85); status_box.info(random.choice(TIER_5_MESSAGES))
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
