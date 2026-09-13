
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
        "review": "",
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


def build_prompt(cm_text: str, up_text: str, topic: str, session: str, lesson_date: str, learning_competency: str, learning_objectives: list[str]) -> str:
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
Topic: {topic}
Session: {session or "[not supplied]"}
Date: {lesson_date or "[not supplied]"}
Learning Competency supplied by teacher: {learning_competency or "[not supplied]"}
Learning Objectives supplied by teacher:
1. {learning_objectives[0] or "[not supplied]"}
2. {learning_objectives[1] or "[not supplied]"}
3. {learning_objectives[2] or "[not supplied]"}

RULES
=====
1. Generate ONE complete daily lesson plan.
2. Match the requested topic to the correct lesson/section in the Curriculum Map.
3. Preserve source terminology, competencies, assessment names, activities, resources,
   values/integrations, references, grade level, subject, term, and unit.
4. Use the Unit Plan's Transfer Goal, Essential Understanding, and Essential Question.
5. Do not borrow topic-specific content from another lesson.
6. Do not invent school-specific facts, staff names, references, or unrelated curriculum requirements.
   Exception: if the Learning Competency is not supplied by the teacher and cannot be clearly found
   in the Curriculum Map or Unit Plan, create one reasonable, topic-aligned competency.
7. Produce EXACTLY THREE learning objectives under lesson_development.presentation_of_concept.
   Each objective must be measurable, topic-specific, and aligned with the competency.
   If the teacher supplies an objective, preserve it. If an objective field is blank, derive a suitable
   objective from the Curriculum Map and Unit Plan; if the source files do not state one, create an
   appropriate objective for the topic.
8. If the teacher supplies a Learning Competency, use it exactly. Otherwise, first derive the competency
   from the Curriculum Map and Unit Plan, and only create one if it is not stated there.
9. Activities may be expanded into classroom-ready instructions and steps, but must preserve
   the purpose of the activity named in the Curriculum Map or Unit Plan.
10. Formative and summative assessments must match the source documents.
11. Keep the lesson realistic for one daily session.
12. The Action statement must begin with "I will..."
13. Preserve the Biblical reference in the Curriculum Map. Do not invent a direct Bible quotation
    if the actual verse wording is not provided by the uploaded files. In that case, leave "text" blank.
14. Use the user-supplied Session and Date when present; otherwise leave them blank.
15. Use simple, natural teacher wording. Write as a classroom teacher would write a daily lesson plan.
    Keep sentences clear, practical, and concise. Avoid overly formal, flowery, technical, or AI-sounding wording.
16. Keep directions and guide questions short and easy to understand. Do not add long explanations unless the
    Curriculum Map or Unit Plan specifically requires them.
17. Return ONLY valid JSON. Do not use markdown fences or add explanations.

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
    cm_text: str,
    up_text: str,
    topic: str,
    session: str,
    lesson_date: str,
    learning_competency: str,
    learning_objectives: list[str],
) -> Dict[str, Any]:
    if not GEMINI_API_KEY:
        raise RuntimeError(
            "The website owner has not configured GEMINI_API_KEY in the server secrets."
        )

    client = genai.Client(api_key=GEMINI_API_KEY)
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=build_prompt(
            cm_text,
            up_text,
            topic,
            session,
            lesson_date,
            learning_competency,
            learning_objectives,
        ),
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
    Build the learning plan using the school's actual Word layout.
    A4, Arial 12, same bordered sections, signature boxes, logo, and implementation area.
    """
    if not TEMPLATE_PATH.exists():
        raise RuntimeError(
            "The school Word template file is missing. "
            "Please upload school_learning_plan_template.docx to the GitHub repository."
        )

    doc = Document(str(TEMPLATE_PATH))

    # Required school page setup.
    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.top_margin = Inches(.5)
    sec.bottom_margin = Inches(.5)
    sec.left_margin = Inches(.5)
    sec.right_margin = Inches(.5)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    normal.font.size = Pt(12)

    # Keep the school's heading/logo area from the template.
    # Make sure all existing heading text remains Arial 12.
    for p in doc.paragraphs:
        for run in p.runs:
            _school_set_run_font(run, bold=run.bold, italic=run.italic)

    # --------------------------------------------------------
    # TOP INFORMATION TABLE
    # --------------------------------------------------------
    meta = doc.tables[0]

    row0 = [
        ("Subject:", d.get("subject", "")),
        ("Term:", d.get("term", "")),
        ("Session:", d.get("session", "")),
    ]
    for i, (label, value) in enumerate(row0):
        label_cell = meta.rows[0].cells[i * 2]
        value_cell = meta.rows[0].cells[i * 2 + 1]
        _school_clear_cell(label_cell)
        _school_clear_cell(value_cell)
        _school_add_p(label_cell, label, bold=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        _school_add_p(value_cell, value, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    row1 = [
        ("Topic:", d.get("topic", "")),
        ("Unit:", d.get("unit", "")),
    ]
    for i, (label, value) in enumerate(row1):
        label_cell = meta.rows[1].cells[i * 2]
        value_cell = meta.rows[1].cells[i * 2 + 1]
        _school_clear_cell(label_cell)
        _school_clear_cell(value_cell)
        _school_add_p(label_cell, label, bold=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        _school_add_p(value_cell, value, align=WD_ALIGN_PARAGRAPH.LEFT)

    date_cell = meta.rows[1].cells[4]
    _school_clear_cell(date_cell)
    _school_add_p(date_cell, "Date:", bold=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    date_value = str(d.get("date", "") or "").strip()
    if date_value:
        for line_value in date_value.splitlines():
            if line_value.strip():
                _school_add_p(
                    date_cell,
                    line_value.strip(),
                    left=.25,
                    align=WD_ALIGN_PARAGRAPH.LEFT,
                )

    competency_cell = meta.rows[2].cells[0]
    _school_clear_cell(competency_cell)
    _school_add_p(
        competency_cell,
        "Learning Competencies:",
        bold=True,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )
    competencies = d.get("learning_competencies", []) or []
    for i, competency in enumerate(competencies):
        if str(competency).strip():
            _school_add_p(
                competency_cell,
                f"{chr(97 + i)}.  {str(competency).strip()}",
                left=.25,
                align=WD_ALIGN_PARAGRAPH.LEFT,
            )

    # --------------------------------------------------------
    # MAIN SCHOOL FORMAT TABLE
    # --------------------------------------------------------
    table = doc.tables[1]

    # Transfer Goal / Essential Understanding / Essential Question
    fixed_sections = [
        ("Transfer Goal", "transfer_goal"),
        ("Essential Understanding", "essential_understanding"),
        ("Essential Question", "essential_question"),
    ]
    for row_index, (heading, key) in enumerate(fixed_sections):
        cell = table.rows[row_index].cells[0]
        _school_clear_cell(cell)
        _school_add_p(
            cell,
            heading,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )
        _school_add_p(
            cell,
            d.get(key, ""),
            align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        )

    # I. Preliminaries
    cell = table.rows[3].cells[0]
    _school_clear_cell(cell)
    prelim = d.get("preliminaries", {}) or {}

    _school_add_p(cell, "I.  Preliminaries", bold=True, left=.02, align=WD_ALIGN_PARAGRAPH.LEFT)

    _school_add_p(cell, "A.  Review", bold=True, left=.30, align=WD_ALIGN_PARAGRAPH.LEFT)
    _school_add_p(cell, prelim.get("review", ""), left=.55)

    _school_add_p(cell, "B.  Focus", bold=True, left=.30, align=WD_ALIGN_PARAGRAPH.LEFT)
    _school_add_p(cell, prelim.get("focus", ""), left=.55, align=WD_ALIGN_PARAGRAPH.LEFT)

    _school_add_p(cell, "C.  Resources", bold=True, left=.30, align=WD_ALIGN_PARAGRAPH.LEFT)
    for resource in prelim.get("resources", []) or []:
        if str(resource).strip():
            _school_add_p(
                cell,
                f"•  {str(resource).strip()}",
                left=.55,
                align=WD_ALIGN_PARAGRAPH.LEFT,
            )

    _school_add_p(cell, "D.  Motivation", bold=True, left=.30, align=WD_ALIGN_PARAGRAPH.LEFT)
    motivation = prelim.get("motivation", {}) or {}
    if motivation.get("title"):
        _school_add_p(
            cell,
            motivation.get("title", ""),
            bold=True,
            italic=True,
            left=.55,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )
    if motivation.get("instruction"):
        _school_add_label_value(
            cell,
            "Instruction: ",
            motivation.get("instruction", ""),
            left=.55,
        )
    if motivation.get("content"):
        _school_add_p(
            cell,
            motivation.get("content", ""),
            left=.55,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )

    motivation_questions = motivation.get("guide_questions", []) or []
    if motivation_questions:
        guide_label = "Guide Question:" if len(motivation_questions) == 1 else "Guide Questions:"
        _school_add_p(
            cell,
            guide_label,
            bold=True,
            italic=True,
            left=.55,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )
        for i, question in enumerate(motivation_questions, 1):
            prefix = "" if len(motivation_questions) == 1 else f"{i}.) "
            _school_add_p(cell, prefix + str(question), left=.55)

    _school_add_p(
        cell,
        "E.  Activating Prior Knowledge",
        bold=True,
        left=.30,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )
    _school_add_p(
        cell,
        prelim.get("activating_prior_knowledge", ""),
        left=.55,
    )

    # II. Lesson Development
    cell = table.rows[4].cells[0]
    _school_clear_cell(cell)
    lesson = d.get("lesson_development", {}) or {}

    _school_add_p(cell, "II.  Lesson Development", bold=True, left=.02, align=WD_ALIGN_PARAGRAPH.LEFT)

    _school_add_p(
        cell,
        "A.  Presentation of Concept",
        bold=True,
        left=.30,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )
    _school_add_p(
        cell,
        "The students will be able to…",
        left=.55,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )

    # The generator is designed to produce exactly three learning objectives.
    objectives = lesson.get("presentation_of_concept", []) or []
    for i, objective in enumerate(objectives[:3]):
        if str(objective).strip():
            _school_add_p(
                cell,
                f"{chr(97 + i)}.  {str(objective).strip()}",
                left=.55,
            )

    _school_add_p(cell, "B.  Activities", bold=True, left=.30, align=WD_ALIGN_PARAGRAPH.LEFT)
    for activity in lesson.get("activities", []) or []:
        activity_type = str(activity.get("type", "") or "").strip()
        activity_title = str(activity.get("title", "") or "").strip()
        activity_prefix = f"{activity_type} Activity: " if activity_type else "Activity: "

        _school_add_p(
            cell,
            activity_prefix + activity_title,
            bold=True,
            left=.55,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )

        if activity.get("instruction"):
            _school_add_label_value(
                cell,
                "Instruction: ",
                activity.get("instruction", ""),
                left=.55,
            )

        steps = activity.get("steps", []) or []
        if steps:
            _school_add_p(
                cell,
                "Procedure:",
                bold=True,
                left=.55,
                align=WD_ALIGN_PARAGRAPH.LEFT,
            )
            for i, step in enumerate(steps, 1):
                _school_add_p(cell, f"{i}.) {step}", left=.55)

        activity_questions = activity.get("guide_questions", []) or []
        if activity_questions:
            label = "Guide Question:" if len(activity_questions) == 1 else "Guide Questions:"
            _school_add_p(
                cell,
                label,
                bold=True,
                italic=True,
                left=.55,
                align=WD_ALIGN_PARAGRAPH.LEFT,
            )
            for i, question in enumerate(activity_questions, 1):
                _school_add_p(cell, f"{i}.) {question}", left=.55)

    _school_add_p(
        cell,
        "C.  Broadening of Concept",
        bold=True,
        left=.30,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )
    broadening = lesson.get("broadening_of_concept", {}) or {}
    broadening_items = [
        ("Leading Question", "leading_question"),
        ("Exploring Question", "exploring_question"),
        ("Connecting Question", "connecting_question"),
        ("Essential Question", "essential_question"),
    ]
    for label, key in broadening_items:
        if broadening.get(key):
            _school_add_p(
                cell,
                f"•  {label}",
                bold=True,
                left=.55,
                align=WD_ALIGN_PARAGRAPH.LEFT,
            )
            _school_add_p(cell, broadening.get(key, ""), left=.55)

    _school_add_p(cell, "D.  Integration", bold=True, left=.30, align=WD_ALIGN_PARAGRAPH.LEFT)
    integration = lesson.get("integration", {}) or {}

    integrations = [
        ("Ignacian Core Value", "ignacian_core_value", "connection"),
        ("Related Core Value", "related_core_value", "connection"),
        ("Social Orientation", "social_orientation", "question_or_connection"),
        ("Lesson Across Discipline", "lesson_across_discipline", "question_or_connection"),
    ]
    for label, key, detail_key in integrations:
        item = integration.get(key, {}) or {}
        if item.get("name"):
            _school_add_label_value(
                cell,
                f"•  {label}: ",
                item.get("name", ""),
                left=.55,
            )
            if item.get(detail_key):
                _school_add_p(cell, item.get(detail_key, ""), left=.55)

    biblical = integration.get("biblical_text_reflection", {}) or {}
    if biblical.get("reference"):
        _school_add_label_value(
            cell,
            "•  Biblical Text/Reflection: ",
            biblical.get("reference", ""),
            left=.55,
        )
    if biblical.get("text"):
        _school_add_p(cell, biblical.get("text", ""), left=.55)
    if biblical.get("reflection"):
        _school_add_p(cell, biblical.get("reflection", ""), left=.55)

    # III. Evaluation/Assessment
    cell = table.rows[5].cells[0]
    _school_clear_cell(cell)
    evaluation = d.get("evaluation_assessment", {}) or {}

    _school_add_p(
        cell,
        "III.  Evaluation/Assessment",
        bold=True,
        left=.02,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )
    _school_add_p(
        cell,
        "A.  Formative Assessment",
        bold=True,
        left=.30,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )
    for item in evaluation.get("formative", []) or []:
        if item.get("name"):
            value = f"•  {item.get('name', '')}"
            if item.get("instruction"):
                value += f": {item.get('instruction', '')}"
            _school_add_p(cell, value, left=.55)

    _school_add_p(
        cell,
        "B.  Summative Assessment",
        bold=True,
        left=.30,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )
    for item in evaluation.get("summative", []) or []:
        if item.get("name"):
            value = f"•  {item.get('name', '')}"
            if item.get("instruction"):
                value += f": {item.get('instruction', '')}"
            _school_add_p(cell, value, left=.55)

    # IV. Summary/Action
    cell = table.rows[6].cells[0]
    _school_clear_cell(cell)
    summary_action = d.get("summary_action", {}) or {}

    _school_add_p(cell, "IV.  Summary/Action", bold=True, left=.02, align=WD_ALIGN_PARAGRAPH.LEFT)
    _school_add_p(cell, "A.  Summary", bold=True, left=.30, align=WD_ALIGN_PARAGRAPH.LEFT)
    _school_add_p(cell, summary_action.get("summary", ""), left=.55)
    _school_add_p(cell, "B.  Action", bold=True, left=.30, align=WD_ALIGN_PARAGRAPH.LEFT)
    _school_add_p(cell, summary_action.get("action", ""), left=.55)

    # V. Purposive Assignment/Enrichment
    cell = table.rows[7].cells[0]
    _school_clear_cell(cell)
    assignment = d.get("assignment_enrichment", {}) or {}

    _school_add_p(
        cell,
        "V.  Purposive Assignment/Enrichment",
        bold=True,
        left=.02,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )
    if assignment.get("assignment"):
        _school_add_label_value(
            cell,
            "Assignment: ",
            assignment.get("assignment", ""),
            left=.30,
        )
    if assignment.get("instructions"):
        _school_add_label_value(
            cell,
            "Instruction: ",
            assignment.get("instructions", ""),
            left=.30,
        )

    assignment_questions = assignment.get("guide_questions", []) or []
    if assignment_questions:
        label = "Guide Question:" if len(assignment_questions) == 1 else "Guide Questions:"
        _school_add_p(
            cell,
            label,
            bold=True,
            left=.30,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )
        for i, question in enumerate(assignment_questions, 1):
            _school_add_p(cell, f"{i}.) {question}", left=.55)

    # VI. References
    cell = table.rows[8].cells[0]
    _school_clear_cell(cell)
    _school_add_p(cell, "VI.  References", bold=True, left=.02, align=WD_ALIGN_PARAGRAPH.LEFT)
    for reference in d.get("references", []) or []:
        if str(reference).strip():
            _school_add_p(cell, str(reference).strip(), left=.30)

    # --------------------------------------------------------
    # SIGNATURES - same boxed layout as the school sample
    # --------------------------------------------------------
    prepared_title = str(d.get("prepared_by_title", "") or "").strip()
    if not prepared_title:
        subject_text = str(d.get("subject", "") or "").strip()
        prepared_title = f"{subject_text} Teacher" if subject_text else "Teacher"

    signature_row = table.rows[9]
    signature_data = [
        (
            signature_row.cells[0],
            "Prepared by:",
            d.get("prepared_by", ""),
            prepared_title,
        ),
        (
            signature_row.cells[2],
            "Checked by:",
            d.get("checked_by", ""),
            d.get("checked_by_title", ""),
        ),
    ]
    for cell, label, name, title in signature_data:
        _school_clear_cell(cell)
        _school_add_p(cell, label, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        _school_add_p(cell, "", align=WD_ALIGN_PARAGRAPH.LEFT)
        _school_add_p(cell, name, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _school_add_p(cell, title, align=WD_ALIGN_PARAGRAPH.CENTER)

    noted_cell = table.rows[10].cells[1]
    _school_clear_cell(noted_cell)
    _school_add_p(noted_cell, "Noted by:", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    _school_add_p(noted_cell, "", align=WD_ALIGN_PARAGRAPH.LEFT)
    _school_add_p(
        noted_cell,
        d.get("noted_by", ""),
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    _school_add_p(
        noted_cell,
        d.get("noted_by_title", ""),
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )

    # --------------------------------------------------------
    # STATUS OF IMPLEMENTATION / MODIFICATIONS
    # --------------------------------------------------------
    status_row = table.rows[11]

    status_cell = status_row.cells[0]
    _school_clear_cell(status_cell)
    _school_add_p(
        status_cell,
        "Status of Implementation",
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    _school_add_p(
        status_cell,
        "___Implemented         ___ Partially Implemented",
        left=.05,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )
    _school_add_p(
        status_cell,
        "___Not Implemented",
        left=.05,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )
    _school_add_p(status_cell, "Remarks:", bold=True, left=.05, align=WD_ALIGN_PARAGRAPH.LEFT)
    _school_add_p(status_cell, "_____________________________________", left=.05, align=WD_ALIGN_PARAGRAPH.LEFT)
    _school_add_p(status_cell, "_____________________________________", left=.05, align=WD_ALIGN_PARAGRAPH.LEFT)
    _school_add_p(status_cell, "", align=WD_ALIGN_PARAGRAPH.LEFT)
    _school_add_p(
        status_cell,
        "Observed by: _________________________",
        left=.05,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )
    _school_add_p(status_cell, "", align=WD_ALIGN_PARAGRAPH.LEFT)
    _school_add_p(
        status_cell,
        "Date Observed: _______________________",
        left=.05,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )

    modifications_cell = status_row.cells[2]
    _school_clear_cell(modifications_cell)
    _school_add_p(
        modifications_cell,
        "Modifications",
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    for _ in range(3):
        _school_add_p(
            modifications_cell,
            "_____________________________________",
            left=.05,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )
    _school_add_p(modifications_cell, "", align=WD_ALIGN_PARAGRAPH.LEFT)
    _school_add_p(
        modifications_cell,
        "Remarks:",
        bold=True,
        left=.05,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )
    for _ in range(3):
        _school_add_p(
            modifications_cell,
            "_____________________________________",
            left=.05,
            align=WD_ALIGN_PARAGRAPH.LEFT,
        )

    out = io.BytesIO()
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
c_teacher, c_grade = st.columns(2)
with c_teacher:
    teacher_name = st.text_input(
        "Teacher Name *",
        placeholder="Example: Maria Santos"
    )
with c_grade:
    grade_level_input = st.text_input(
        "Grade Level *",
        placeholder="Example: Grade 8"
    )

learning_competency = st.text_area(
    "Learning Competency (Optional)",
    placeholder="Leave blank to let AI derive it from the Curriculum Map and Unit Plan.",
    height=90
)

st.markdown("**Learning Objectives (Optional)**")
st.caption("You may enter up to three objectives. Any blank objective will be completed by AI based on the Curriculum Map and Unit Plan.")

obj1 = st.text_input(
    "Learning Objective 1 (Optional)",
    placeholder="Leave blank for AI-generated objective."
)
obj2 = st.text_input(
    "Learning Objective 2 (Optional)",
    placeholder="Leave blank for AI-generated objective."
)
obj3 = st.text_input(
    "Learning Objective 3 (Optional)",
    placeholder="Leave blank for AI-generated objective."
)

topic = st.text_input(
    " Topic *",
    placeholder="Example: Factoring"
)

c3, c4 = st.columns(2)
with c3:
    session = st.text_input("Session", placeholder="Example: 6")
with c4:
    lesson_date = st.text_input("Date", placeholder="Example: September 15, 2026")

st.subheader("3. Requirements Check")

ready = True
if cm_file:
    st.markdown('<div class="status-ok">Curriculum Map uploaded</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="status-no">Curriculum Map required</div>', unsafe_allow_html=True)
    ready = False

if up_file:
    st.markdown('<div class="status-ok">Unit Plan uploaded</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="status-no">Unit Plan required</div>', unsafe_allow_html=True)
    ready = False

if teacher_name.strip():
    st.markdown(
        f'<div class="status-ok">Teacher: {teacher_name.strip()}</div>',
        unsafe_allow_html=True
    )
else:
    st.markdown('<div class="status-no">Teacher Name required</div>', unsafe_allow_html=True)
    ready = False

if grade_level_input.strip():
    st.markdown(
        f'<div class="status-ok">Grade Level: {grade_level_input.strip()}</div>',
        unsafe_allow_html=True
    )
else:
    st.markdown('<div class="status-no">Grade Level required</div>', unsafe_allow_html=True)
    ready = False

if topic.strip():
    st.markdown(f'<div class="status-ok">Topic: {topic.strip()}</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="status-no">Topic required</div>', unsafe_allow_html=True)
    ready = False

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
            user_objectives = [obj1.strip(), obj2.strip(), obj3.strip()]

            data = generate_plan(
                cm_text=cm_text,
                up_text=up_text,
                topic=topic.strip(),
                session=session.strip(),
                lesson_date=lesson_date.strip(),
                learning_competency=learning_competency.strip(),
                learning_objectives=user_objectives,
            )

            # Required form values are inserted directly into the Word document data.
            data["grade_level"] = grade_level_input.strip()
            data["prepared_by"] = teacher_name.strip()

            # Learning Competency is optional:
            # use the teacher's entry when supplied; otherwise keep the AI-derived/generated value.
            if learning_competency.strip():
                data["learning_competencies"] = [learning_competency.strip()]

            # Learning Objectives are optional:
            # preserve teacher-entered objectives and let AI fill blank positions.
            lesson_dev = data.setdefault("lesson_development", {})
            ai_objectives = lesson_dev.get("presentation_of_concept", [])
            if not isinstance(ai_objectives, list):
                ai_objectives = [str(ai_objectives)] if ai_objectives else []

            merged_objectives = []
            for i in range(3):
                if user_objectives[i]:
                    merged_objectives.append(user_objectives[i])
                elif i < len(ai_objectives) and str(ai_objectives[i]).strip():
                    merged_objectives.append(str(ai_objectives[i]).strip())
                else:
                    merged_objectives.append("")

            lesson_dev["presentation_of_concept"] = merged_objectives

        with st.spinner("Preparing the Word document..."):
            docx_bytes = build_docx(data)

        st.session_state["generated_docx"] = docx_bytes
        st.session_state["generated_topic"] = topic.strip()
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
