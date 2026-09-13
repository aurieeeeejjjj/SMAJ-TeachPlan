
import io
import json
import os
import re
from typing import Any, Dict, List

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
from openai import OpenAI


# ============================================================
# APP CONFIG
# ============================================================

st.set_page_config(
    page_title="Daily Lesson Plan Generator",
    page_icon="📘",
    layout="centered",
)

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
    <h1>📘 Daily Lesson Plan Generator</h1>
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

OPENAI_API_KEY = get_secret("OPENAI_API_KEY")
OPENAI_MODEL = get_secret("OPENAI_MODEL", "gpt-5.6-luna")
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


def build_prompt(cm_text: str, up_text: str, topic: str, session: str, lesson_date: str) -> str:
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

RULES
=====
1. Generate ONE complete daily lesson plan.
2. Match the requested topic to the correct lesson/section in the Curriculum Map.
3. Preserve source terminology, competencies, assessment names, activities, resources,
   values/integrations, references, grade level, subject, term, and unit.
4. Use the Unit Plan's Transfer Goal, Essential Understanding, and Essential Question.
5. Do not borrow topic-specific content from another lesson.
6. Do not invent school-specific facts, staff names, references, or curriculum requirements.
   If a field is not supported by the uploaded files, leave it blank.
7. Presentation objectives must be measurable, topic-specific, and aligned with the competency.
8. Activities may be expanded into classroom-ready instructions and steps, but must preserve
   the purpose of the activity named in the Curriculum Map or Unit Plan.
9. Formative and summative assessments must match the source documents.
10. Keep the lesson realistic for one daily session.
11. The Action statement must begin with "I will..."
12. Preserve the Biblical reference in the Curriculum Map. Do not invent a direct Bible quotation
    if the actual verse wording is not provided by the uploaded files. In that case, leave "text" blank.
13. Use the user-supplied Session and Date when present; otherwise leave them blank.
14. Return ONLY valid JSON. Do not use markdown fences or add explanations.

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
) -> Dict[str, Any]:
    if not OPENAI_API_KEY:
        raise RuntimeError(
            "The website owner has not configured OPENAI_API_KEY in the server secrets."
        )

    client = OpenAI(api_key=OPENAI_API_KEY)
    response = client.responses.create(
        model=OPENAI_MODEL,
        input=build_prompt(cm_text, up_text, topic, session, lesson_date),
    )
    return extract_json(response.output_text)


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


def build_docx(d: Dict[str, Any]) -> bytes:
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(.45)
    sec.bottom_margin = Inches(.45)
    sec.left_margin = Inches(.55)
    sec.right_margin = Inches(.55)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    normal.font.size = Pt(10.5)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    add_text(p, d.get("school_name", ""), bold=True, size=13)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    school_year = d.get("school_year", "")
    add_text(p, f"School Year {school_year}".strip(), size=11)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(5)
    add_text(p, "LEARNING PLAN", bold=True, size=13)

    meta = doc.add_table(rows=4, cols=4)
    rows = [
        ("Subject:", d.get("subject",""), "Grade Level:", d.get("grade_level","")),
        ("Term:", d.get("term",""), "Unit:", d.get("unit","")),
        ("Topic:", d.get("topic",""), "Session:", d.get("session","")),
        ("Date:", d.get("date",""), "", "")
    ]
    for r, vals in enumerate(rows):
        for c, val in enumerate(vals):
            p = meta.rows[r].cells[c].paragraphs[0]
            add_text(p, val, bold=(c % 2 == 0 and bool(val)))

    section_heading(doc, "Learning Competencies:")
    bullets(doc, d.get("learning_competencies", []))

    section_heading(doc, "Transfer Goal")
    line(doc, value=d.get("transfer_goal",""), indent=0)

    section_heading(doc, "Essential Understanding")
    line(doc, value=d.get("essential_understanding",""), indent=0)

    section_heading(doc, "Essential Question")
    line(doc, value=d.get("essential_question",""), indent=0)

    # I
    pr = d.get("preliminaries", {})
    section_heading(doc, "I. Preliminaries")

    subheading(doc, "A. Review")
    line(doc, value=pr.get("review",""))

    subheading(doc, "B. Focus")
    line(doc, value=pr.get("focus",""))

    subheading(doc, "C. Resources")
    bullets(doc, pr.get("resources", []))

    subheading(doc, "D. Motivation")
    mot = pr.get("motivation", {})
    if mot.get("title"):
        line(doc, value=mot.get("title",""))
    if mot.get("instruction"):
        line(doc, "Instruction: ", mot.get("instruction",""))
    if mot.get("content"):
        line(doc, value=mot.get("content",""))
    if mot.get("guide_questions"):
        line(doc, "Guide Question/s:")
        numbered(doc, mot.get("guide_questions", []))

    subheading(doc, "E. Activating Prior Knowledge")
    line(doc, value=pr.get("activating_prior_knowledge",""))

    # II
    ld = d.get("lesson_development", {})
    section_heading(doc, "II. Lesson Development")

    subheading(doc, "A. Presentation of Concept")
    line(doc, value="The students will be able to…")
    for i, obj in enumerate(ld.get("presentation_of_concept", []), 1):
        if obj:
            letter = chr(96 + i) if i <= 26 else str(i)
            line(doc, value=f"{letter}. {obj}", indent=.55)

    subheading(doc, "B. Activities")
    for act in ld.get("activities", []):
        title = act.get("title","")
        typ = act.get("type","")
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(.35)
        if typ:
            add_text(p, f"{typ} Activity: ", bold=True)
        else:
            add_text(p, "Activity: ", bold=True)
        add_text(p, title, bold=True)

        if act.get("instruction"):
            line(doc, "Instruction: ", act.get("instruction",""), indent=.45)
        if act.get("steps"):
            line(doc, "Procedure:", indent=.45)
            numbered(doc, act.get("steps", []), indent=.65)
        if act.get("guide_questions"):
            line(doc, "Guide Question/s:", indent=.45)
            numbered(doc, act.get("guide_questions", []), indent=.65)

    subheading(doc, "C. Broadening of Concept")
    broad = ld.get("broadening_of_concept", {})
    line(doc, "• Leading Question: ", broad.get("leading_question",""))
    line(doc, "• Exploring Question: ", broad.get("exploring_question",""))
    line(doc, "• Connecting Question: ", broad.get("connecting_question",""))
    line(doc, "• Essential Question: ", broad.get("essential_question",""))

    subheading(doc, "D. Integration")
    integ = ld.get("integration", {})

    iv = integ.get("ignacian_core_value", {})
    line(doc, "• Ignacian Core Value: ", iv.get("name",""))
    if iv.get("connection"):
        line(doc, value=iv.get("connection",""), indent=.55)

    rv = integ.get("related_core_value", {})
    line(doc, "• Related Core Value: ", rv.get("name",""))
    if rv.get("connection"):
        line(doc, value=rv.get("connection",""), indent=.55)

    so = integ.get("social_orientation", {})
    line(doc, "• Social Orientation: ", so.get("name",""))
    if so.get("question_or_connection"):
        line(doc, value=so.get("question_or_connection",""), indent=.55)

    lad = integ.get("lesson_across_discipline", {})
    line(doc, "• Lesson Across Discipline: ", lad.get("name",""))
    if lad.get("question_or_connection"):
        line(doc, value=lad.get("question_or_connection",""), indent=.55)

    bible = integ.get("biblical_text_reflection", {})
    line(doc, "• Biblical Text/Reflection: ", bible.get("reference",""))
    if bible.get("text"):
        line(doc, value=bible.get("text",""), indent=.55)
    if bible.get("reflection"):
        line(doc, value=bible.get("reflection",""), indent=.55)

    # III
    eva = d.get("evaluation_assessment", {})
    section_heading(doc, "III. Evaluation/Assessment")

    subheading(doc, "A. Formative Assessment")
    for item in eva.get("formative", []):
        if item.get("name"):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(.35)
            add_text(p, "• " + item.get("name",""), bold=True)
            if item.get("instruction"):
                add_text(p, ": " + item.get("instruction",""))

    subheading(doc, "B. Summative Assessment")
    for item in eva.get("summative", []):
        if item.get("name"):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(.35)
            add_text(p, "• " + item.get("name",""), bold=True)
            if item.get("instruction"):
                add_text(p, ": " + item.get("instruction",""))

    # IV
    sa = d.get("summary_action", {})
    section_heading(doc, "IV. Summary/Action")
    subheading(doc, "A. Summary")
    line(doc, value=sa.get("summary",""))
    subheading(doc, "B. Action")
    line(doc, value=sa.get("action",""))

    # V
    ae = d.get("assignment_enrichment", {})
    section_heading(doc, "V. Purposive Assignment/Enrichment")
    if ae.get("assignment"):
        line(doc, "Assignment: ", ae.get("assignment",""))
    if ae.get("instructions"):
        line(doc, "Instruction: ", ae.get("instructions",""))
    if ae.get("guide_questions"):
        line(doc, "Guide Question/s:")
        numbered(doc, ae.get("guide_questions", []))

    # VI
    section_heading(doc, "VI. References")
    for ref in d.get("references", []):
        if ref:
            line(doc, value=ref, indent=0)

    # Signatures
    doc.add_paragraph()
    sig = doc.add_table(rows=1, cols=3)
    sig_items = [
        ("Prepared by:", d.get("prepared_by",""), d.get("prepared_by_title","")),
        ("Checked by:", d.get("checked_by",""), d.get("checked_by_title","")),
        ("Noted by:", d.get("noted_by",""), d.get("noted_by_title","")),
    ]
    for i, (label, name, title) in enumerate(sig_items):
        cell = sig.rows[0].cells[i]
        p = cell.paragraphs[0]
        add_text(p, label, bold=True)
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_text(p, name, bold=True)
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_text(p, title)

    section_heading(doc, "Status of Implementation")
    line(doc, value="___ Implemented     ___ Partially Implemented     ___ Not Implemented", indent=0)
    line(doc, "Remarks:")
    line(doc, value="____________________________________________")
    line(doc, value="____________________________________________")
    line(doc, value="Observed by: ________________________________")
    line(doc, value="Date Observed: ______________________________")

    section_heading(doc, "Modifications")
    line(doc, value="____________________________________________")
    line(doc, value="____________________________________________")
    line(doc, value="____________________________________________")
    line(doc, "Remarks:")
    line(doc, value="____________________________________________")
    line(doc, value="____________________________________________")

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
        "📑 Curriculum Map (.docx)",
        type=["docx"],
        key="curriculum_map"
    )

with c2:
    up_file = st.file_uploader(
        "📘 Unit Plan (.docx)",
        type=["docx"],
        key="unit_plan"
    )

st.subheader("2. Lesson Details")
topic = st.text_input(
    "✏️ Topic *",
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
    st.markdown('<div class="status-ok">✓ Curriculum Map uploaded</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="status-no">✗ Curriculum Map required</div>', unsafe_allow_html=True)
    ready = False

if up_file:
    st.markdown('<div class="status-ok">✓ Unit Plan uploaded</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="status-no">✗ Unit Plan required</div>', unsafe_allow_html=True)
    ready = False

if topic.strip():
    st.markdown(f'<div class="status-ok">✓ Topic: {topic.strip()}</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="status-no">✗ Topic required</div>', unsafe_allow_html=True)
    ready = False

if not OPENAI_API_KEY:
    st.warning(
        "Website setup is incomplete: the owner still needs to add the server API key."
    )

st.divider()

if st.button(
    "✨ GENERATE LESSON PLAN",
    type="primary",
    disabled=(not ready or not OPENAI_API_KEY),
):
    try:
        with st.spinner("Reading and aligning your curriculum documents..."):
            cm_text = read_docx(cm_file)
            up_text = read_docx(up_file)

        with st.spinner("Creating your Daily Learning Plan..."):
            data = generate_plan(
                cm_text=cm_text,
                up_text=up_text,
                topic=topic.strip(),
                session=session.strip(),
                lesson_date=lesson_date.strip(),
            )

        with st.spinner("Preparing the Word document..."):
            docx_bytes = build_docx(data)

        st.session_state["generated_docx"] = docx_bytes
        st.session_state["generated_topic"] = topic.strip()
        st.success("✅ Lesson plan ready!")

    except Exception as exc:
        st.error("The lesson plan could not be generated.")
        st.exception(exc)

if "generated_docx" in st.session_state:
    safe_topic = re.sub(
        r"[^A-Za-z0-9_-]+",
        "_",
        st.session_state.get("generated_topic", "Lesson")
    ).strip("_") or "Lesson"

    st.download_button(
        "📄 DOWNLOAD WORD FILE (.DOCX)",
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
