
import io
import json
import os
import re
from typing import Any, Dict, List
from pathlib import Path

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, Mm
from docx.oxml.ns import qn
from google import genai
from google.genai import types
from io import BytesIO


TEMPLATE_PATH = Path(__file__).with_name("school_learning_plan_template.docx")


# ============================================================
# APP CONFIG
# ============================================================

st.set_page_config(
    page_title="Daily Lesson Plan Generator",
    page_icon="",
    layout="centered",
)

st.markdown("""
<style>
:root {
    --purple-main: #7C3AED;
    --purple-dark: #5B21B6;
    --purple-soft: #F3E8FF;
    --purple-pale: #FAF5FF;
    --ink: #24123A;
}
.stApp {
    background: linear-gradient(180deg, #FFFFFF 0%, #FCFAFF 42%, #FAF5FF 100%);
}
h1, h2, h3 {
    color: var(--ink);
}
div[data-testid="stFileUploader"] {
    background: #FAF5FF;
    border-radius: 16px;
}
div[data-testid="stFileUploaderDropzone"] {
    background: #F3E8FF;
    border: 1px dashed #A78BFA;
    border-radius: 14px;
}
div[data-testid="stTextInput"] input {
    background: #FCFAFF;
    border: 1px solid #D8B4FE;
    border-radius: 10px;
}
div[data-testid="stTextInput"] input:focus {
    border-color: #7C3AED;
    box-shadow: 0 0 0 1px #7C3AED;
}
.stButton > button {
    background: linear-gradient(90deg, #7C3AED, #6D28D9);
    color: white;
    border: none;
    border-radius: 12px;
    font-weight: 700;
    min-height: 3rem;
}
.stButton > button:hover {
    background: #5B21B6;
    color: white;
    border: none;
}
.stDownloadButton > button {
    background: #7C3AED;
    color: white;
    border: none;
    border-radius: 12px;
    font-weight: 700;
}
.stDownloadButton > button:hover {
    background: #5B21B6;
    color: white;
}
.status-ok {
    background: #F3E8FF !important;
    border-left: 4px solid #7C3AED !important;
    color: #4C1D95 !important;
}
.status-no {
    border-radius: 10px;
}
</style>
""", unsafe_allow_html=True)


st.markdown("""
<style>
.block-container {
    max-width: 900px;
    padding-top: 2rem;
    padding-bottom: 3rem;
}
.hero {
    text-align: center;
    padding: 1rem 0 .6rem 0;
}
.hero h1 {
    margin-bottom: .2rem;
    font-size: 2.15rem;
}
.hero p {
    opacity: .72;
    margin-top: 0;
}
.section-card {
    border: 1px solid rgba(127,127,127,.22);
    border-radius: 16px;
    padding: 1.1rem 1.15rem;
    margin-bottom: 1rem;
}
div.stButton > button, div.stDownloadButton > button {
    width: 100%;
    border-radius: 12px;
    min-height: 3rem;
    font-weight: 700;
}
.status-ok {
    padding: .6rem .8rem;
    border-radius: 10px;
    background: rgba(0,180,100,.09);
    margin-bottom: .45rem;
}
.status-no {
    padding: .6rem .8rem;
    border-radius: 10px;
    background: rgba(220,50,50,.08);
    margin-bottom: .45rem;
}
.small {
    font-size: .9rem;
    opacity: .72;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <h1> Daily Lesson Plan Generator</h1>
    <p>Upload your Curriculum Map and Unit Plan, enter the topic, and download a Word lesson plan.</p>
</div>
""", unsafe_allow_html=True)


# ============================================================
# SECRET / ENV SETTINGS
# ============================================================

def get_secret(name: str, default: str = "") -> str:
    try:
        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:
        pass
    return os.getenv(name, default)

GEMINI_API_KEY = get_secret("GEMINI_API_KEY")
GEMINI_MODEL = get_secret("GEMINI_MODEL", "gemini-2.5-flash-lite")
SITE_PASSWORD = get_secret("SITE_PASSWORD", "")


# ============================================================
# OPTIONAL SITE PASSWORD
# ============================================================

if SITE_PASSWORD:
    if "site_unlocked" not in st.session_state:
        st.session_state.site_unlocked = False

    if not st.session_state.site_unlocked:
        st.subheader("Teacher Access")
        pw = st.text_input("Enter site password", type="password")
        if st.button("Open Generator"):
            if pw == SITE_PASSWORD:
                st.session_state.site_unlocked = True
                st.rerun()
            else:
                st.error("Incorrect password.")
        st.stop()


# ============================================================
# DOCX INPUT READER
# ============================================================

def read_docx(uploaded_file) -> str:
    doc = Document(uploaded_file)
    parts: List[str] = []

    for p in doc.paragraphs:
        text = p.text.strip()
        if text:
            parts.append(text)

    for table in doc.tables:
        for row in table.rows:
            row_text = []
            for cell in row.cells:
                cell_text = " ".join(
                    p.text.strip()
                    for p in cell.paragraphs
                    if p.text.strip()
                )
                if cell_text:
                    row_text.append(cell_text)
            if row_text:
                parts.append(" | ".join(row_text))

    return "\n".join(parts)


# ============================================================
# STRUCTURED OUTPUT SCHEMA
# ============================================================

SCHEMA: Dict[str, Any] = {
    "school_name": "",
    "school_year": "",
    "subject": "",
    "grade_level": "",
    "term": "",
    "session": "",
    "topic": "",
    "unit": "",
    "date": "",
    "learning_competencies": [""],
    "transfer_goal": "",
    "essential_understanding": "",
    "essential_question": "",
    "preliminaries": {
        "review": {"lots_question": "", "hots_question": ""},
        "focus": "",
        "resources": [""],
        "motivation": {
            "title": "",
            "instruction": "",
            "content": "",
            "guide_questions": [""]
        },
        "activating_prior_knowledge": ""
    },
    "lesson_development": {
        "presentation_of_concept": [""],
        "activities": [
            {
                "title": "",
                "type": "",
                "instruction": "",
                "steps": [""],
                "guide_questions": [""]
            }
        ],
        "broadening_of_concept": {
            "leading_question": "",
            "exploring_question": "",
            "connecting_question": "",
            "essential_question": ""
        },
        "integration": {
            "ignacian_core_value": {"name": "", "connection": ""},
            "related_core_value": {"name": "", "connection": ""},
            "social_orientation": {"name": "", "question_or_connection": ""},
            "lesson_across_discipline": {"name": "", "question_or_connection": ""},
            "biblical_text_reflection": {"reference": "", "text": "", "reflection": ""}
        }
    },
    "evaluation_assessment": {
        "formative": [{"name": "", "instruction": ""}],
        "summative": [{"name": "", "instruction": ""}]
    },
    "summary_action": {
        "summary": "",
        "action": ""
    },
    "assignment_enrichment": {
        "assignment": "",
        "instructions": "",
        "guide_questions": [""]
    },
    "references": [""],
    "prepared_by": "",
    "prepared_by_title": "",
    "checked_by": "",
    "checked_by_title": "",
    "noted_by": "",
    "noted_by_title": ""
}


def build_prompt(cm_text: str, up_text: str, curriculum_topic: str, specific_lesson_focus: str, session: str, lesson_date: str, language: str, subject: str, grade_sections: str, term: str, customization: str) -> str:
    return f"""
You are a Daily Learning Plan generator for teachers.

Use ONLY the supplied Curriculum Map and Unit Plan as the authoritative basis for
school-specific and lesson-specific content.

CURRICULUM MAP
==============
{cm_text}

UNIT PLAN
=========
{up_text}

REQUESTED LESSON
================
Language: {language}
Subject: {subject}
Grade and Section/s: {grade_sections}
Term: {term}
Topic from Curriculum Map: {curriculum_topic}
Specific Lesson Topic / Focus: {specific_lesson_focus}
Session: {session or "[not supplied]"}
Date: {lesson_date or "[not supplied]"}
Teacher Customization / Contextualization: {customization or "[none supplied]"}

RULES
=====
1. Generate ONE realistic daily lesson plan for the Specific Lesson Topic / Focus in the selected Language (English or Filipino).
2. Use simple, natural teacher wording: short, clear, practical, understandable sentences, directions, and questions.
3. Respect Customization / Contextualization unless it conflicts with the Curriculum Map, Unit Plan, or competency.
4. Keep the COMPLETE lesson realistic for LESS THAN ONE HOUR.
5. Use ONLY Learning Competency/Competencies supported by the Curriculum Map or Unit Plan. NEVER invent one. If none matches, return an empty learning_competencies list.
6. Strictly align: Learning Competency -> Specific Lesson Focus -> Daily Objectives -> Presentation -> Activities -> Formative Assessment -> Summative Assessment -> Summary -> Action.
7. Objectives are based on the competency and MUST NOT exceed its cognitive/performance demand. If it says "discuss," do not require "demonstrate" unless supported.
8. Transfer Goal, Essential Understanding, and Essential Question MUST come from the relevant Unit Plan when available.
9. Review uses the previous lesson in source sequence and contains EXACTLY TWO questions: 1 LOTS and 1 HOTS, connecting previous learning to today's lesson.
10. Focus is today's Specific Lesson Topic / Focus.
11. Resources use source-listed resources first; add only genuinely needed practical resources such as TV, PPT Presentation, HDMI, Textbook, Notebook, Paper, and Pen.
12. Motivation directly connects to the lesson, catches attention, is engaging, and brief.
13. Activating Prior Knowledge is a simple starter QUESTION needed for today's lesson.
14. Presentation of Concept supplies measurable daily objectives introduced in Word by "The students will be able to…"; never exceed the competency.
15. Activities align with competency, objectives, Curriculum Map, and focus. Prefer source activities when suitable; otherwise create a simple aligned activity. Use group/individual/both only when appropriate and feasible under one hour.
16. Broadening provides one simple lesson-related question for EACH Leading, Exploring, Connecting, and Essential Question, progressively deepening thinking.
17. Ignacian Core Value is Faith, Excellence, or Service. Follow source if stated; otherwise choose the natural fit.
18. Related values: FAITH—Strong Faith in God, Prophetic Witness to Gospel Values, Nationalism, Justice, Communion. EXCELLENCE—Integrity, Competence, Resourcefulness, Discipline, Self-reliance. SERVICE—Stewardship, Humility, Charity, Courage, Preferential Love of the Poor.
19. Social Orientation connects learning to family, school, community, or society.
20. Lesson Across Discipline meaningfully connects another subject with a simple question/connection.
21. Biblical Text/Reflection genuinely aligns with the lesson. Prefer source material and NEVER fabricate verse wording.
22. Formative Assessment happens DURING lesson/activity and checks competency/objectives.
23. Summative Assessment checks today's learning near the end and measures the same competency/objectives.
24. Summary is a QUESTION leading students to summarize/explain/apply the main learning.
25. Action is a QUESTION leading students to apply learning in a real situation.
26. Purposive Assignment/Enrichment strengthens today's learning or prepares tomorrow's lesson.
27. References use Unit Plan/Curriculum Map first. Added content/ideas require their source. Use APA style as available details allow; never invent missing bibliographic details.
28. Do not invent staff names, unsupported curriculum requirements, or school-specific facts.
29. Use teacher-supplied Subject, Grade and Section/s, Term, Session, and Date where appropriate.
30. Return ONLY valid JSON.

RETURN EXACTLY THIS JSON SHAPE
==============================
{json.dumps(SCHEMA, indent=2)}
"""


def extract_json(text: str) -> Dict[str, Any]:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start >= 0 and end > start:
            return json.loads(text[start:end+1])
        raise


def generate_plan(
    cm_text: str, up_text: str, curriculum_topic: str, specific_lesson_focus: str,
    session: str, lesson_date: str, language: str, subject: str,
    grade_sections: str, term: str, customization: str,
) -> Dict[str, Any]:
    if not GEMINI_API_KEY:
        raise RuntimeError(
            "The website owner has not configured GEMINI_API_KEY in the server secrets."
        )

    client = genai.Client(api_key=GEMINI_API_KEY)
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=build_prompt(cm_text, up_text, curriculum_topic, specific_lesson_focus, session, lesson_date, language, subject, grade_sections, term, customization),
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.2,
            max_output_tokens=12000,
        ),
    )

    if not response.text:
        raise RuntimeError("Gemini returned an empty response. Please try again.")

    return extract_json(response.text)


# ============================================================
# WORD DOCUMENT BUILDER
# ============================================================

def set_font(run, size=10.5, bold=False, italic=False):
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


def add_text(p, text, *, bold=False, italic=False, size=10.5):
    r = p.add_run(text or "")
    set_font(r, size=size, bold=bold, italic=italic)
    return r


def section_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    add_text(p, text, bold=True, size=11)
    return p


def subheading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(.18)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(1)
    add_text(p, text, bold=True)
    return p


def line(doc, label="", value="", indent=.35):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(indent)
    p.paragraph_format.space_after = Pt(2)
    if label:
        add_text(p, label, bold=True)
    if value:
        add_text(p, value)
    return p


def bullets(doc, items, indent=.48):
    for item in items or []:
        if not item:
            continue
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.left_indent = Inches(indent)
        p.paragraph_format.space_after = Pt(1)
        add_text(p, item)


def numbered(doc, items, indent=.55):
    for i, item in enumerate(items or [], 1):
        if not item:
            continue
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(indent)
        p.paragraph_format.first_line_indent = Inches(-.2)
        p.paragraph_format.space_after = Pt(1)
        add_text(p, f"{i}.) {item}")


def _school_set_run_font(run, *, bold=None, italic=None):
    """Apply the school's required Word font."""
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    run.font.size = Pt(12)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def _school_clear_cell(cell):
    """Clear only the contents while preserving the school's table borders/size."""
    cell._tc.clear_content()


def _school_add_p(
    cell,
    text="",
    *,
    bold=False,
    italic=False,
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    left=0,
):
    p = cell.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    if left:
        pf.left_indent = Inches(left)

    r = p.add_run(str(text or ""))
    _school_set_run_font(r, bold=bold, italic=italic)
    return p


def _school_add_label_value(cell, label, value, *, left=.35, italic_value=False):
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    if left:
        pf.left_indent = Inches(left)

    r = p.add_run(label)
    _school_set_run_font(r, bold=True)
    r = p.add_run(str(value or ""))
    _school_set_run_font(r, italic=italic_value)
    return p


def build_docx(d: Dict[str, Any]) -> bytes:
    """
    Populate the FINAL official SMAJ Learning Plan template.
    English uses the original English headings.
    Filipino keeps the exact same Word layout and translates the headings.
    """
    if not TEMPLATE_PATH.exists():
        raise RuntimeError(
            "The final school Word template is missing. "
            "Please keep school_learning_plan_template.docx beside app.py."
        )

    doc = Document(str(TEMPLATE_PATH))

    # Preserve the final template's page/layout structure, while enforcing A4 + Arial 12.
    for sec in doc.sections:
        sec.page_width = Mm(210)
        sec.page_height = Mm(297)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    normal.font.size = Pt(12)

    for p in doc.paragraphs:
        for run in p.runs:
            _school_set_run_font(run, bold=run.bold, italic=run.italic)

    lang = str(d.get("language", "English") or "English").strip().lower()
    fil = lang == "filipino"

    L = {
        "plan": "BANGHAY-ARALIN" if fil else "LEARNING PLAN",
        "subject_level": "Asignatura/Baitang: " if fil else "Subject/Level: ",
        "unit": "Yunit: " if fil else "Unit: ",
        "topic": "Paksa: " if fil else "Topic: ",
        "term": "Markahan/Termino: " if fil else "Quarter/Term: ",
        "session": "Sesyon: " if fil else "Session: ",
        "date": "Petsa: " if fil else "Date: ",
        "section": "Seksyon: " if fil else "Section: ",
        "competency": "Mga Kasanayang Pampagkatuto:" if fil else "Learning Competency (s):",
        "transfer": "LAYUNIN SA PAGLILIPAT NG PAGKATUTO" if fil else "TRANSFER GOAL",
        "understanding": "MAHALAGANG PAG-UNAWA" if fil else "ESSENTIAL UNDERSTANDING",
        "essential_q": "MAHAHALAGANG TANONG" if fil else "ESSENTIAL QUESTIONS",
        "prelim": "I. PANIMULANG GAWAIN" if fil else "I. PRELIMINARIES",
        "review": "Balik-Aral" if fil else "Review",
        "lots": "LOTS",
        "hots": "HOTS",
        "focus": "Pokus" if fil else "Focus",
        "resources": "Mga Kagamitan" if fil else "Resources",
        "motivation": "Pagganyak" if fil else "Motivation",
        "instruction": "Panuto: " if fil else "Instructions: ",
        "guide_q": "Gabay na Tanong:" if fil else "Guide Question:",
        "guide_qs": "Mga Gabay na Tanong:" if fil else "Guide Questions:",
        "prior": "Pagpapagana ng Dating Kaalaman" if fil else "Activating Prior Knowledge",
        "development": "II. PAGLINANG NG ARALIN" if fil else "II. LESSON DEVELOPMENT",
        "presentation": "A. Paglalahad ng Konsepto" if fil else "A. Presentation of the Concept",
        "students": "Ang mga mag-aaral ay inaasahang…" if fil else "The students will be able to…",
        "activities": "B. Mga Gawain" if fil else "B. Activities",
        "individual": "Indibidwal na Gawain" if fil else "Individual Activity",
        "group": "Pangkatang Gawain" if fil else "Group Activity",
        "activity": "Gawain" if fil else "Activity",
        "procedure": "Pamamaraan:" if fil else "Procedure:",
        "broadening": "C. Pagpapalawak ng Konsepto" if fil else "C. Broadening of Concept",
        "leading": "Panimulang Tanong" if fil else "Leading Question",
        "exploring": "Mapanuring Tanong" if fil else "Exploring Question",
        "connecting": "Tanong na Nag-uugnay" if fil else "Connecting Question",
        "essential": "Mahalagang Tanong" if fil else "Essential Question",
        "integration": "D. Integrasyon" if fil else "D. Integration",
        "ignacian": "Pangunahing Pagpapahalagang Ignacian" if fil else "Ignacian Core Value",
        "related": "Kaugnay na Pagpapahalaga" if fil else "Related Value",
        "social": "Oryentasyong Panlipunan" if fil else "Social Orientation",
        "discipline": "Ugnayan sa Ibang Disiplina" if fil else "Lesson Across Discipline",
        "biblical": "Tekstong Biblikal/Pagninilay" if fil else "Biblical Text/Reflection",
        "evaluation": "III. PAGTATAYA" if fil else "III. EVALUATION/ASSESSMENT",
        "formative": "PORMATIBONG PAGTATAYA" if fil else "FORMATIVE ASSESSMENT",
        "summative": "SUMATIBONG PAGTATAYA" if fil else "SUMMATIVE ASSESSMENT",
        "summary_action": "IV. PAGLALAGOM/PAGKILOS" if fil else "IV. SUMMARY/ACTION",
        "summary": "Paglalagom" if fil else "Summary",
        "action": "Pagkilos" if fil else "Action",
        "assignment": "V. MAKABULUHANG TAKDANG-ARAL/PAGPAPAYAMAN" if fil else "V. PURPOSIVE ASSIGNMENT/ENRICHMENT",
        "references": "VI. MGA SANGGUNIAN" if fil else "VI. REFERENCES",
        "prepared": "Inihanda ni:" if fil else "Prepared by:",
        "teacher": "Guro" if fil else "Teacher",
        "submitted": "Petsa ng Pagsumite:_________________" if fil else "Date Submitted:_________________",
        "checked": "Sinuri ni:" if fil else "Checked by:",
        "noted": "Pinagtibay ni:" if fil else "Noted by:",
        "status": "Kalagayan ng Pagpapatupad" if fil else "Status of Implementation",
        "implemented": "___ Naipatupad     ___ Bahagyang Naipatupad" if fil else "___ Implemented   ___ Partially Implemented",
        "not_implemented": "___ Hindi Naipatupad" if fil else "___ Not Implemented",
        "remarks": "Mga Tala:" if fil else "Remarks:",
        "observed": "Inobserbahan ni:_________________________" if fil else "Observed by:_________________________",
        "date_observed": "Petsa ng Obserbasyon:____________________" if fil else "Date Observed:_______________________",
        "modifications": "Mga Pagbabago" if fil else "Modifications",
    }

    def clear(cell):
        _school_clear_cell(cell)

    def add(cell, text="", bold=False, italic=False, left=0, align=WD_ALIGN_PARAGRAPH.LEFT):
        return _school_add_p(cell, text, bold=bold, italic=italic, left=left, align=align)

    def add_lv(cell, label, value, left=0):
        return _school_add_label_value(cell, label, value, left=left)

    # Translate only the title line; preserve the official school header/logo.
    for p in doc.paragraphs:
        if p.text.strip().upper() == "LEARNING PLAN":
            for r in p.runs:
                r.text = ""
            r = p.add_run(L["plan"])
            _school_set_run_font(r, bold=True)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            break

    # ---------------- FINAL TOP TABLE: exact 3 x 3 layout ----------------
    meta = doc.tables[0]

    grade_sections_text = str(d.get("grade_level", "") or "").strip()
    grade_match = re.search(r"(?i)grade\s*(\d+)", grade_sections_text)
    level_text = grade_match.group(1) if grade_match else ""
    subject_level_text = f'{d.get("subject","")} {level_text}'.strip()

    vals = [
        (0, 0, L["subject_level"], subject_level_text),
        (0, 1, L["unit"], d.get("unit","")),
        (0, 2, L["topic"], d.get("topic","")),
        (1, 0, L["term"], d.get("term","")),
        (1, 1, L["session"], d.get("session","")),
    ]
    for ri, ci, label, value in vals:
        c = meta.rows[ri].cells[ci]
        clear(c)
        add_lv(c, label, value, left=0)

    c = meta.rows[1].cells[2]
    clear(c)
    add_lv(c, L["date"], d.get("date",""), left=0)
    add_lv(c, L["section"], d.get("grade_level",""), left=0)

    c = meta.rows[2].cells[0]
    clear(c)
    add(c, L["competency"], bold=True)
    for i, comp in enumerate(d.get("learning_competencies",[]) or [], 1):
        if str(comp).strip():
            add(c, f"{i}. {str(comp).strip()}", left=.05)

    # ---------------- FINAL MAIN TABLE: exact 13-row layout ----------------
    table = doc.tables[1]

    for ri, heading, key in [
        (0, L["transfer"], "transfer_goal"),
        (1, L["understanding"], "essential_understanding"),
        (2, L["essential_q"], "essential_question"),
    ]:
        c = table.rows[ri].cells[0]
        clear(c)
        add(c, heading, bold=True)
        add(c, d.get(key,""), align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    prelim = d.get("preliminaries",{}) or {}
    c = table.rows[3].cells[0]
    clear(c)
    add(c, L["prelim"], bold=True)
    add(c, L["review"], bold=True)
    review = prelim.get("review",{}) or {}
    if isinstance(review,dict):
        if str(review.get("lots_question","") or "").strip():
            add_lv(c, f'{L["lots"]}: ', review.get("lots_question",""))
        if str(review.get("hots_question","") or "").strip():
            add_lv(c, f'{L["hots"]}: ', review.get("hots_question",""))
    elif str(review).strip():
        add(c, str(review).strip())

    add(c, L["focus"], bold=True)
    add(c, d.get("topic", prelim.get("focus","")))

    add(c, L["resources"], bold=True)
    for r in prelim.get("resources",[]) or []:
        if str(r).strip(): add(c, f"o  {str(r).strip()}", left=.15)

    add(c, L["motivation"], bold=True)
    mot = prelim.get("motivation",{}) or {}
    if mot.get("title"): add(c, mot.get("title",""), bold=True)
    if mot.get("instruction"): add_lv(c, L["instruction"], mot.get("instruction",""))
    if mot.get("content"): add(c, mot.get("content",""))
    qs = mot.get("guide_questions",[]) or []
    if qs:
        add(c, L["guide_q"] if len(qs)==1 else L["guide_qs"], bold=True)
        for q in qs: add(c, f"o  {q}", left=.15)

    add(c, L["prior"], bold=True)
    prior = prelim.get("activating_prior_knowledge","")
    if isinstance(prior,list):
        for q in prior:
            if str(q).strip(): add(c, f"o  {q}", left=.15)
    elif str(prior).strip():
        add(c, str(prior).strip())

    lesson = d.get("lesson_development",{}) or {}
    c = table.rows[4].cells[0]
    clear(c)
    add(c, L["development"], bold=True)
    add(c, L["presentation"], bold=True)
    add(c, L["students"], italic=True)
    for i,obj in enumerate(lesson.get("presentation_of_concept",[]) or [],1):
        if str(obj).strip(): add(c, f"{i}. {str(obj).strip()}")

    add(c, L["activities"], bold=True)
    type_counts = {}
    for act in lesson.get("activities",[]) or []:
        typ = str(act.get("type","") or "").strip()
        typ_low = typ.lower()
        if "group" in typ_low or "pangkat" in typ_low:
            type_label = L["group"]
        elif "individual" in typ_low or "indibid" in typ_low:
            type_label = L["individual"]
        else:
            type_label = typ or L["activity"]
        type_counts[type_label] = type_counts.get(type_label,0)+1
        add(c, type_label, bold=True)
        title = str(act.get("title","") or "").strip()
        if title: add(c, f'{L["activity"]} {type_counts[type_label]}: {title}', bold=True)
        if act.get("instruction"): add_lv(c, L["instruction"], act.get("instruction",""))
        steps = act.get("steps",[]) or []
        if steps:
            add(c, L["procedure"], bold=True)
            for step in steps: add(c, f"o  {step}", left=.15)
        aqs = act.get("guide_questions",[]) or []
        if aqs:
            add(c, L["guide_q"] if len(aqs)==1 else L["guide_qs"], bold=True)
            for q in aqs: add(c, f"o  {q}", left=.15)

    add(c, L["broadening"], bold=True)
    broad = lesson.get("broadening_of_concept",{}) or {}
    for label,key in [
        (L["leading"],"leading_question"),
        (L["exploring"],"exploring_question"),
        (L["connecting"],"connecting_question"),
        (L["essential"],"essential_question"),
    ]:
        if broad.get(key): add_lv(c, f"{label}: ", broad.get(key,""))

    c = table.rows[5].cells[0]
    clear(c)
    add(c, L["integration"], bold=True)
    integ = lesson.get("integration",{}) or {}
    for label,key,detail in [
        (L["ignacian"],"ignacian_core_value","connection"),
        (L["related"],"related_core_value","connection"),
        (L["social"],"social_orientation","question_or_connection"),
        (L["discipline"],"lesson_across_discipline","question_or_connection"),
    ]:
        item = integ.get(key,{}) or {}
        if item.get("name"):
            add_lv(c, f"{label}: ", item.get("name",""))
        if item.get(detail): add(c, item.get(detail,""))

    bib = integ.get("biblical_text_reflection",{}) or {}
    if bib.get("reference"): add_lv(c, f'{L["biblical"]}: ', bib.get("reference",""))
    if bib.get("text"): add(c, bib.get("text",""), italic=True)
    if bib.get("reflection"): add(c, bib.get("reflection",""))

    ev = d.get("evaluation_assessment",{}) or {}
    c = table.rows[6].cells[0]
    clear(c)
    add(c, L["evaluation"], bold=True)
    add(c, L["formative"], bold=True)
    for item in ev.get("formative",[]) or []:
        name = str(item.get("name","") or "").strip()
        ins = str(item.get("instruction","") or "").strip()
        if name or ins: add(c, "o  " + (name + (": " if name and ins else "") + ins))
    add(c, L["summative"], bold=True)
    for item in ev.get("summative",[]) or []:
        name = str(item.get("name","") or "").strip()
        ins = str(item.get("instruction","") or "").strip()
        if name or ins: add(c, "o  " + (name + (": " if name and ins else "") + ins))

    sa = d.get("summary_action",{}) or {}
    c = table.rows[7].cells[0]
    clear(c)
    add(c, L["summary_action"], bold=True)
    add(c, L["summary"], bold=True)
    add(c, sa.get("summary",""))
    add(c, L["action"], bold=True)
    add(c, sa.get("action",""))

    ass = d.get("assignment_enrichment",{}) or {}
    c = table.rows[8].cells[0]
    clear(c)
    add(c, L["assignment"], bold=True)
    if ass.get("assignment"): add(c, ass.get("assignment",""))
    if ass.get("instructions"): add_lv(c, L["instruction"], ass.get("instructions",""))
    for q in ass.get("guide_questions",[]) or []:
        add(c, f"o  {q}", left=.15)

    c = table.rows[9].cells[0]
    clear(c)
    add(c, L["references"], bold=True)
    for ref in d.get("references",[]) or []:
        if str(ref).strip(): add(c, str(ref).strip())

    # Signatures: preserve exact two-column / noted / implementation layout.
    left = table.rows[10].cells[0]
    right = table.rows[10].cells[1]
    clear(left); clear(right)
    add(left, L["prepared"], bold=True)
    add(left, "")
    add(left, d.get("prepared_by",""), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    prepared_title = str(d.get("prepared_by_title","") or "").strip() or L["teacher"]
    add(left, prepared_title, align=WD_ALIGN_PARAGRAPH.CENTER)
    add(left, "")
    add(left, L["submitted"])

    add(right, L["checked"], bold=True)
    add(right, "")
    add(right, d.get("checked_by",""), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add(right, d.get("checked_by_title",""), align=WD_ALIGN_PARAGRAPH.CENTER)

    noted = table.rows[11].cells[0]
    clear(noted)
    add(noted, L["noted"], bold=True)
    add(noted, "")
    add(noted, d.get("noted_by",""), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add(noted, d.get("noted_by_title",""), align=WD_ALIGN_PARAGRAPH.CENTER)

    status = table.rows[12].cells[0]
    mods = table.rows[12].cells[1]
    clear(status); clear(mods)
    add(status, L["status"], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add(status, L["implemented"])
    add(status, L["not_implemented"])
    add(status, L["remarks"], bold=True)
    add(status, "____________________________________")
    add(status, L["observed"])
    add(status, L["date_observed"])

    add(mods, L["modifications"], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add(mods, "")
    add(mods, "__________________________________________________________________________________________")
    add(mods, L["remarks"], bold=True)
    add(mods, "")
    add(mods, "")

    # Remove the unused blank body paragraphs that follow the final table in the supplied
    # template. They otherwise create an extra blank page after shorter generated plans.
    body = doc._element.body
    last_table_seen = False
    for child in list(body):
        tag = child.tag.split("}")[-1]
        if tag == "tbl":
            last_table_seen = True
            continue
        if last_table_seen and tag == "p":
            text_nodes = child.xpath(".//w:t/text()")
            if not "".join(text_nodes).strip():
                body.remove(child)

    # Enforce Arial 12 for all generated/existing text without changing the template structure.
    for table_obj in doc.tables:
        for row in table_obj.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        _school_set_run_font(run, bold=run.bold, italic=run.italic)

    out = BytesIO()
    doc.save(out)
    return out.getvalue()

# ============================================================
# USER INTERFACE
# ============================================================

st.subheader("1. Upload Required Documents")
c1, c2 = st.columns(2)

with c1:
    cm_file = st.file_uploader(
        " Curriculum Map (.docx)",
        type=["docx"],
        key="curriculum_map"
    )

with c2:
    up_file = st.file_uploader(
        " Unit Plan (.docx)",
        type=["docx"],
        key="unit_plan"
    )

st.subheader("2. Teacher & Lesson Details")
c_teacher, c_subject = st.columns(2)
with c_teacher:
    teacher_name = st.text_input("Teacher's Name *", placeholder="Example: Maria Santos")
with c_subject:
    subject_input = st.text_input("Subject *", placeholder="Example: Mathematics")
c_grade, c_term = st.columns(2)
with c_grade:
    grade_sections_input = st.text_input("Grade and Section/s *", placeholder="Example: Grade 8 - St. Luke, St. John")
with c_term:
    term_input = st.text_input("Term *", placeholder="Example: First Term")
language = st.selectbox("Language *", ["English", "Filipino"])
curriculum_topic = st.text_input("Topic from Curriculum Map *", placeholder="Example: Measures of Central Tendency")
specific_lesson_focus = st.text_input("Specific Lesson Topic / Focus *", placeholder="Example: Finding the Mean")
st.caption("Narrow a broad curriculum topic into the lesson for the day. The same curriculum topic may be used for several daily lesson plans.")
customization = st.text_area("Customization / Contextualization (Optional)", placeholder="Add special instructions, local examples, preferred activities, student context, or anything you want adjusted.", height=110)
c3, c4 = st.columns(2)
with c3: session = st.text_input("Session (Optional)", placeholder="Example: Session 1")
with c4: lesson_date = st.text_input("Date (Optional)", placeholder="Example: September 15, 2026")

st.subheader("3. Requirements Check")

ready = True
for uploaded, ok_text, no_text in [(cm_file,"Curriculum Map uploaded","Curriculum Map required"),(up_file,"Unit Plan uploaded","Unit Plan required")]:
    if uploaded: st.markdown(f'<div class="status-ok">{ok_text}</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="status-no">{no_text}</div>', unsafe_allow_html=True); ready=False
for label,value in [("Teacher",teacher_name),("Subject",subject_input),("Grade and Section/s",grade_sections_input),("Term",term_input),("Curriculum Topic",curriculum_topic),("Lesson Focus",specific_lesson_focus)]:
    if value.strip(): st.markdown(f'<div class="status-ok">{label}: {value.strip()}</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="status-no">{label} required</div>', unsafe_allow_html=True); ready=False
st.markdown(f'<div class="status-ok">Language: {language}</div>', unsafe_allow_html=True)

if not GEMINI_API_KEY:
    st.warning(
        "Website setup is incomplete: the owner still needs to add the server API key."
    )

st.divider()

if st.button(
    " GENERATE LESSON PLAN",
    type="primary",
    disabled=(not ready or not GEMINI_API_KEY),
):
    try:
        with st.spinner("Reading and aligning your curriculum documents..."):
            cm_text = read_docx(cm_file)
            up_text = read_docx(up_file)

        with st.spinner("Creating your Daily Learning Plan..."):
            data = generate_plan(
                cm_text=cm_text, up_text=up_text, curriculum_topic=curriculum_topic.strip(),
                specific_lesson_focus=specific_lesson_focus.strip(), session=session.strip(),
                lesson_date=lesson_date.strip(), language=language, subject=subject_input.strip(),
                grade_sections=grade_sections_input.strip(), term=term_input.strip(), customization=customization.strip(),
            )
            data["language"]=language
            data["subject"]=subject_input.strip()
            data["grade_level"]=grade_sections_input.strip()
            data["term"]=term_input.strip()
            data["prepared_by"]=teacher_name.strip()
            data["topic"]=specific_lesson_focus.strip()

            competencies = data.get("learning_competencies", []) or []
            competencies = [str(c).strip() for c in competencies if str(c).strip()]
            data["learning_competencies"] = competencies

            if not competencies:
                st.warning(
                    "No matching Learning Competency was found in the uploaded Curriculum Map or Unit Plan "
                    "for this topic and lesson focus. Please check the topic/focus and try again."
                )
                st.stop()

        with st.spinner("Preparing the Word document..."):
            docx_bytes = build_docx(data)

        st.session_state["generated_docx"] = docx_bytes
        st.session_state["generated_topic"] = specific_lesson_focus.strip()
        st.success(" Lesson plan ready!")

    except Exception as exc:
        raw_error = str(exc)
        error_text = raw_error.lower()

        if "429" in error_text or "quota" in error_text or "resource_exhausted" in error_text or "rate" in error_text:
            st.error(
                "The Gemini free-tier limit has been reached for now. Please wait and try again later."
            )
        elif (
            "api key" in error_text
            or "api_key" in error_text
            or "permission" in error_text
            or "unauthenticated" in error_text
            or "401" in error_text
            or "403" in error_text
        ):
            st.error(
                "Gemini could not authenticate this app. Please check the GEMINI_API_KEY saved in Streamlit Secrets."
            )
        elif "404" in error_text or "not found" in error_text or "model" in error_text and "not" in error_text:
            st.error(
                "The configured Gemini model is unavailable. Check GEMINI_MODEL in Streamlit Secrets."
            )
        elif "json" in error_text or "decode" in error_text:
            st.error(
                "Gemini returned an incomplete lesson plan response. Please click Generate Lesson Plan again."
            )
        else:
            st.error(
                "The lesson plan could not be generated. The app owner can use the diagnostic message below."
            )
            # Show only a short diagnostic message. API keys are redacted if they ever appear.
            safe_error = re.sub(r'AIza[0-9A-Za-z_-]{20,}', '[REDACTED_API_KEY]', raw_error)
            safe_error = safe_error[:800]
            st.code(f"{type(exc).__name__}: {safe_error}")

        print(f"Lesson plan generation error ({type(exc).__name__}): {raw_error}")

if "generated_docx" in st.session_state:
    safe_topic = re.sub(
        r"[^A-Za-z0-9_-]+",
        "_",
        st.session_state.get("generated_topic", "Lesson")
    ).strip("_") or "Lesson"

    st.download_button(
        " DOWNLOAD WORD FILE (.DOCX)",
        data=st.session_state["generated_docx"],
        file_name=f"Daily_Learning_Plan_{safe_topic}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
    )

st.markdown(
    '<p class="small" style="text-align:center;margin-top:1.4rem;">'
    'Your uploaded files are used to generate the requested lesson plan during this session.'
    '</p>',
    unsafe_allow_html=True
)
