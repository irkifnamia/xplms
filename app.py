"""XPLMS — Experience Point Learning Management System."""

from __future__ import annotations

import io
import hashlib
import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import bcrypt
import pandas as pd
import plotly.express as px
import streamlit as st
from PIL import Image
from pydantic import BaseModel, Field

try:
    from docx import Document
    from openai import OpenAI
    from pptx import Presentation
    from pypdf import PdfReader
except ImportError:  # Features remain unavailable until requirements are installed.
    Document = OpenAI = Presentation = PdfReader = None

try:
    from supabase import Client, create_client
except ImportError:  # pragma: no cover
    Client = Any
    create_client = None


class GeneratedQuestion(BaseModel):
    question: str
    options: list[str] = Field(min_length=4, max_length=4)
    correct_index: int = Field(ge=0, le=3)
    explanation: str


class GeneratedQuiz(BaseModel):
    title: str
    questions: list[GeneratedQuestion] = Field(min_length=1, max_length=20)


LOGO_PATH = Path(__file__).parent / "assets" / "xplms-logo.jpg"
LOGO_IMAGE = Image.open(LOGO_PATH)

st.set_page_config(
    page_title="XPLMS",
    page_icon=LOGO_IMAGE,
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    [data-testid="stToolbar"],
    [data-testid="stToolbarActions"],
    [data-testid="stStatusWidget"],
    [data-testid="stDecoration"],
    [data-testid="stDeployButton"],
    .stDeployButton,
    #MainMenu {
        display: none !important;
        visibility: hidden !important;
        height: 0 !important;
    }

    [data-testid="stHeader"] {
        display: block !important;
        visibility: visible !important;
        height: 3.5rem !important;
        min-height: 3.5rem !important;
        background: transparent !important;
        pointer-events: none !important;
    }

    [data-testid="stExpandSidebarButton"],
    [data-testid="collapsedControl"] {
        display: flex !important;
        visibility: visible !important;
        position: fixed !important;
        top: 0.75rem !important;
        left: 0.75rem !important;
        z-index: 1000000 !important;
        pointer-events: auto !important;
    }

    [data-testid="stExpandSidebarButton"],
    [data-testid="collapsedControl"] button {
        width: 2.75rem !important;
        height: 2.75rem !important;
        min-height: 2.75rem !important;
        padding: 0.45rem !important;
        border: 1px solid #c4d1df !important;
        border-radius: 12px !important;
        background: #ffffff !important;
        box-shadow: 0 4px 14px rgba(23,35,58,.14) !important;
    }

    [data-testid="stAppViewContainer"] .block-container {
        padding-top: 1.5rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

BRAND = "#1F5DAA"
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@600;700;800&display=swap');
:root { --ink:#17233a; --muted:#5f6b7a; --brand:#1f5daa; --brand-dark:#154b8c; --soft:#f4f7fa; --line:#d3dce6; }
html, body, [class*="css"] { font-family:'DM Sans',sans-serif; color:var(--ink); }
h1,h2,h3 { font-family:'Manrope',sans-serif !important; letter-spacing:-.035em; }
[data-testid="stAppViewContainer"] { background:#f4f7fa; }
[data-testid="stMainBlockContainer"] p,
[data-testid="stMainBlockContainer"] label,
[data-testid="stMainBlockContainer"] [data-testid="stMarkdownContainer"] { color:var(--ink); }
[data-testid="stMainBlockContainer"] [data-testid="stCaptionContainer"] { color:var(--muted); }
[data-testid="stSidebar"] { background:#fff; border-right:1px solid #d3dce6; }
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stCaption { color:#252525 !important; }
[data-testid="stSidebar"] .stRadio label {
  padding:.48rem .55rem; font-weight:600; border-radius:9px; color:#252525 !important;
}
[data-testid="stSidebar"] .stRadio label:hover { background:#edf4fc; }
[data-testid="stSidebar"] .stRadio label p,
[data-testid="stSidebar"] .stRadio label span { color:#252525 !important; }
[data-testid="stSidebar"] hr { border-color:#d3dce6; }
[data-testid="stSidebar"] [data-baseweb="select"] > div {
  background:#fff !important; border-color:#aaa !important; color:#171717 !important;
}
[data-testid="stSidebar"] [data-baseweb="select"] span { color:#171717 !important; }
[data-testid="stSidebar"] [data-baseweb="select"] svg { fill:#171717; }
.block-container { max-width:1440px; padding:1.8rem 2.2rem 4rem; }
.topline { color:var(--brand); font-weight:700; text-transform:uppercase; letter-spacing:.12em; font-size:.72rem; }
.hero { background:linear-gradient(115deg,#fff 0%,#fff 72%,#f2f7fd 100%); color:#17233a;
  border:1px solid #dbe3ec; border-left:6px solid var(--brand); border-radius:18px;
  padding:26px 30px; margin-bottom:20px; box-shadow:0 8px 24px rgba(0,0,0,.055); }
.hero h1 { color:#17233a; margin:.2rem 0 .35rem; font-size:2.15rem; }
.hero p { color:#536174 !important; margin:0; }
.metric-card,.panel { background:white; border:1px solid var(--line); border-radius:18px; padding:18px 20px;
  box-shadow:0 4px 14px rgba(0,0,0,.045); }
.metric-label { color:var(--muted); font-size:.79rem; font-weight:700; text-transform:uppercase; letter-spacing:.06em; }
.metric-value { color:#17233a; font-family:'Manrope'; font-size:1.7rem; font-weight:800; margin:.15rem 0; }
.metric-note { color:var(--muted); font-size:.82rem; }
.pill { display:inline-block; border-radius:999px; padding:4px 9px; font-size:.73rem; font-weight:700;
  color:#174c87; background:#e8f1fc; border:1px solid #b8d2ef; }
.brand { font-family:'Manrope'; font-size:1.35rem; font-weight:800; letter-spacing:-.04em; color:#17233a; }
.brandmark { display:inline-grid; place-items:center; width:29px; height:29px; margin-right:8px; border-radius:9px;
  background:var(--brand); color:white; }
.userbox { background:#f2f7fd; border:1px solid #c8daee; color:#17233a; padding:12px; border-radius:14px; margin:14px 0 8px; }
.userbox b { color:#17233a !important; }
.userbox small { color:#5f6b7a !important; }
.empty { text-align:center; padding:44px 20px; color:var(--muted); background:white; border:1px dashed #ccd0dc; border-radius:18px; }
.stButton > button, .stDownloadButton > button { border-radius:10px; font-weight:700; min-height:2.65rem; color:#171717; border-color:#bdbdb9; background:#fff; }
.stButton > button p, .stDownloadButton > button p { color:#171717 !important; }
.stButton > button:hover, .stDownloadButton > button:hover { color:var(--brand); border-color:var(--brand); }
.stButton > button[kind="primary"] { background:var(--brand) !important; border-color:var(--brand) !important; color:white !important; }
.stButton > button[kind="primary"] p { color:white !important; }
.stButton > button[kind="primary"]:hover { background:var(--brand-dark) !important; border-color:var(--brand-dark) !important; color:white !important; }

/* Segmented filters: white unselected tabs, blue selected tab, readable labels. */
[data-testid="stSegmentedControl"] > div { background:white !important; border:1px solid #b8b8b4 !important; }
[data-testid="stSegmentedControl"] button,
[data-testid="stBaseButton-segmented_control"] {
  background:white !important; border-color:#d4d4d0 !important; color:#171717 !important;
}
[data-testid="stSegmentedControl"] button:hover,
[data-testid="stBaseButton-segmented_control"]:hover {
  background:#edf4fc !important; color:#154b8c !important;
}
[data-testid="stSegmentedControl"] button p,
[data-testid="stSegmentedControl"] button span,
[data-testid="stBaseButton-segmented_control"] p,
[data-testid="stBaseButton-segmented_control"] span { color:#171717 !important; }
[data-testid="stSegmentedControl"] button[aria-pressed="true"],
[data-testid="stBaseButton-segmented_control"][aria-pressed="true"] {
  background:var(--brand) !important; border-color:var(--brand) !important;
}
[data-testid="stSegmentedControl"] button[aria-pressed="true"] p,
[data-testid="stSegmentedControl"] button[aria-pressed="true"] span,
[data-testid="stBaseButton-segmented_control"][aria-pressed="true"] p,
[data-testid="stBaseButton-segmented_control"][aria-pressed="true"] span { color:white !important; }

/* Form controls must stay light even when the OS or Streamlit defaults are dark. */
[data-testid="stFileUploader"] { border-radius:14px; background:white !important; color:#171717 !important; }
[data-testid="stFileUploaderDropzone"],
[data-testid="stFileUploaderDropzone"] > div {
  background:#fafafa !important; border-color:#b8b8b4 !important; color:#171717 !important;
}
[data-testid="stFileUploaderDropzone"] p,
[data-testid="stFileUploaderDropzone"] span,
[data-testid="stFileUploaderDropzone"] small { color:#363636 !important; }
[data-testid="stFileUploaderDropzone"] button {
  background:white !important; border:1px solid #9b9b97 !important; color:#171717 !important;
}
[data-testid="stFileUploaderDropzone"] button p { color:#171717 !important; }
[data-baseweb="input"] > div, [data-baseweb="textarea"] > div,
[data-testid="stTextInput"] > div > div, [data-testid="stTextArea"] > div > div {
  background:white !important; color:#171717 !important;
}
[data-baseweb="input"] input, [data-baseweb="textarea"] textarea,
[data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea {
  background:white !important; color:#171717 !important; -webkit-text-fill-color:#171717 !important;
  border-color:#aaa !important; caret-color:var(--brand);
}
[data-baseweb="input"] input::placeholder, [data-baseweb="textarea"] textarea::placeholder { color:#737373 !important; opacity:1; }
[data-baseweb="tab-list"] { background:#ececea; border-radius:10px; padding:3px; }
[data-testid="stAlert"] { color:#171717; }
[data-testid="stAlert"] p, [data-testid="stAlert"] span { color:#171717 !important; }
[aria-disabled="true"], button:disabled { opacity:1 !important; }
button:disabled, button:disabled p, button:disabled span {
  background:#efefed !important; color:#686868 !important; border-color:#d3d3cf !important;
}
[data-testid="stDataFrame"] { border:1px solid var(--line); border-radius:14px; overflow:hidden; }

/* Final accessibility layer: never rely on hover to reveal wording. */
[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"] { background:#fff !important; color:#171717 !important; }
[data-testid="stHeader"] * { color:#171717 !important; }
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
[data-testid="stMainBlockContainer"] { color:#171717 !important; }
[data-testid="stMainBlockContainer"] label,
[data-testid="stMainBlockContainer"] label *,
[data-testid="stMainBlockContainer"] p,
[data-testid="stMainBlockContainer"] span { text-shadow:none !important; }
[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] *,
[data-testid="stSidebar"] small { color:#5b5b57 !important; opacity:1 !important; }
[data-testid="stImage"] {
  width:100% !important;
  display:flex !important;
  justify-content:center !important;
  margin-left:auto !important;
  margin-right:auto !important;
  text-align:center !important;
}
[data-testid="stImage"] img {
  display:block !important;
  margin-left:auto !important;
  margin-right:auto !important;
}

/* A collapsed preference must never make desktop navigation disappear. */
@media(min-width:701px){
  section[data-testid="stSidebar"] {
    display:block !important;
    visibility:visible !important;
    min-width:280px !important;
    width:280px !important;
    transform:translateX(0) !important;
  }
}

/* Native buttons, form buttons and segmented button groups. */
button[kind="secondary"],
button[kind="tertiary"],
button[kind="segmented_control"],
[data-testid*="BaseButton-secondary"],
[data-testid*="BaseButton-tertiary"],
[data-testid*="BaseButton-segmented"] {
  background:#fff !important;
  color:#171717 !important;
  border-color:#aaa9a4 !important;
}
button[kind="secondary"] *,
button[kind="tertiary"] *,
button[kind="segmented_control"] *,
[data-testid*="BaseButton-secondary"] *,
[data-testid*="BaseButton-tertiary"] *,
[data-testid*="BaseButton-segmented"] * { color:#171717 !important; opacity:1 !important; }
button[kind="primary"],
[data-testid*="BaseButton-primary"] {
  background:#1f5daa !important; color:#fff !important; border-color:#1f5daa !important;
}
button[kind="primary"] *,
[data-testid*="BaseButton-primary"] * { color:#fff !important; opacity:1 !important; }
button[kind="segmented_control"][aria-pressed="true"],
button[kind="segmented_control"][aria-checked="true"],
[data-testid*="BaseButton-segmented"][aria-pressed="true"],
[data-testid*="BaseButton-segmented"][aria-checked="true"] {
  background:#1f5daa !important; border-color:#1f5daa !important;
}
button[kind="segmented_control"][aria-pressed="true"] *,
button[kind="segmented_control"][aria-checked="true"] *,
[data-testid*="BaseButton-segmented"][aria-pressed="true"] *,
[data-testid*="BaseButton-segmented"][aria-checked="true"] * { color:#fff !important; }

/* Tabs and menu wording. */
[data-baseweb="tab-list"] { background:#fff !important; border-bottom:1px solid #d8d8d3; }
[data-baseweb="tab"] { color:#2b2b29 !important; opacity:1 !important; }
[data-baseweb="tab"] * { color:#2b2b29 !important; opacity:1 !important; }
[data-baseweb="tab"][aria-selected="true"],
[data-baseweb="tab"][aria-selected="true"] * { color:#154b8c !important; font-weight:700; }

/* Every field has a light surface and a dark value/placeholder. */
input, textarea, [contenteditable="true"],
[data-baseweb="select"] > div {
  background:#fff !important; color:#171717 !important; -webkit-text-fill-color:#171717 !important;
}
input::placeholder, textarea::placeholder { color:#696965 !important; opacity:1 !important; }
[data-baseweb="select"] *, [role="option"], [role="option"] * {
  color:#171717 !important; opacity:1 !important;
}
[role="listbox"], [role="option"] { background:#fff !important; }

@media(max-width:700px){
  /* Keep Streamlit chrome hidden, but restore the mobile sidebar opener for Admin. */
  [data-testid="stHeader"] {
    display:block !important;
    visibility:visible !important;
    height:3.5rem !important;
    min-height:3.5rem !important;
    background:transparent !important;
  }
  [data-testid="stHeader"] [data-testid="stToolbar"],
  [data-testid="stHeader"] [data-testid="stToolbarActions"],
  [data-testid="stHeader"] [data-testid="stStatusWidget"],
  [data-testid="stHeader"] [data-testid="stDeployButton"] {
    display:none !important;
  }
  [data-testid="collapsedControl"],
  [data-testid="stExpandSidebarButton"] {
    display:flex !important;
    visibility:visible !important;
    position:fixed !important;
    top:.7rem !important;
    left:.7rem !important;
    z-index:1000000 !important;
  }
  [data-testid="collapsedControl"] button,
  [data-testid="stExpandSidebarButton"] {
    width:2.75rem !important;
    height:2.75rem !important;
    min-height:2.75rem !important;
    padding:.45rem !important;
    border:1px solid #c4d1df !important;
    border-radius:12px !important;
    background:#fff !important;
    box-shadow:0 4px 14px rgba(23,35,58,.14) !important;
  }
  [data-testid="collapsedControl"] button *,
  [data-testid="collapsedControl"] svg,
  [data-testid="stExpandSidebarButton"] *,
  [data-testid="stExpandSidebarButton"] svg {
    color:#17233a !important;
    fill:#17233a !important;
  }
  body:has(.student-sidebar-marker) .block-container {
    padding:4rem 0.72rem 4.5rem !important;
  }
  body:has(.student-sidebar-marker) .hero {
    padding:15px 16px;
    margin-bottom:12px;
    border-left-width:4px;
    border-radius:13px;
  }
  body:has(.student-sidebar-marker) .hero h1 {
    font-size:1.42rem;
    line-height:1.18;
  }
  body:has(.student-sidebar-marker) .hero p {
    font-size:.88rem;
    line-height:1.4;
  }
  body:has(.student-sidebar-marker) .topline { font-size:.64rem; }
  body:has(.student-sidebar-marker) .metric-card,
  body:has(.student-sidebar-marker) .panel {
    padding:13px 14px;
    border-radius:13px;
  }
  body:has(.student-sidebar-marker) .metric-value { font-size:1.32rem; }
  body:has(.student-sidebar-marker) [data-testid="stVerticalBlockBorderWrapper"] {
    border-radius:13px !important;
  }
  body:has(.student-sidebar-marker) [data-testid="stPlotlyChart"] {
    background:#fff;
    border:1px solid var(--line);
    border-radius:13px;
    overflow:hidden;
  }
  body:has(.student-sidebar-marker) [data-testid="stDataFrame"] {
    font-size:.78rem;
  }
  body:has(.student-sidebar-marker) button {
    min-height:2.75rem;
    font-size:.88rem;
  }
  body:has(.student-sidebar-marker) [data-testid="stSegmentedControl"] {
    overflow-x:auto;
    scrollbar-width:none;
  }
  body:has(.student-sidebar-marker) [data-testid="stSegmentedControl"] > div {
    min-width:max-content;
  }

  .block-container { padding:4rem .9rem 4.5rem; }
  .hero { padding:20px; border-radius:17px; }
  .hero h1 { font-size:1.55rem; }
  .metric-value { font-size:1.4rem; }
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


DEMO_STUDENTS = pd.DataFrame(
    [
        {"student_id": "S24001", "name": "Alya Sofea", "programme": "Digital Technology", "cohort": "2024", "email": "alya@example.edu"},
        {"student_id": "S24002", "name": "Haziq Amir", "programme": "Digital Technology", "cohort": "2024", "email": "haziq@example.edu"},
        {"student_id": "S24003", "name": "Mei Lin", "programme": "Information Systems", "cohort": "2024", "email": "meilin@example.edu"},
        {"student_id": "S24004", "name": "Arjun Kumar", "programme": "Information Systems", "cohort": "2024", "email": "arjun@example.edu"},
        {"student_id": "S24005", "name": "Nur Iman", "programme": "Digital Technology", "cohort": "2024", "email": "iman@example.edu"},
    ]
)
DEMO_PROGRESS = pd.DataFrame(
    [
        {"student_id": "S24001", "assignment": 87, "pb": 82, "task": 94, "xp": 2460, "badge": 8},
        {"student_id": "S24002", "assignment": 81, "pb": 78, "task": 89, "xp": 2180, "badge": 7},
        {"student_id": "S24003", "assignment": 92, "pb": 91, "task": 86, "xp": 2720, "badge": 10},
        {"student_id": "S24004", "assignment": 74, "pb": 80, "task": 77, "xp": 1840, "badge": 5},
        {"student_id": "S24005", "assignment": 88, "pb": 85, "task": 91, "xp": 2530, "badge": 9},
    ]
)
EMPTY_MATERIALS = pd.DataFrame(
    columns=[
        "id", "title", "course", "chapter", "material_type", "type",
        "file_path", "uploaded_at",
    ]
)


def db() -> Client | None:
    if create_client is None:
        return None
    try:
        service_key = st.secrets.get("SUPABASE_SECRET_KEY") or st.secrets.get(
            "SUPABASE_SERVICE_ROLE_KEY"
        )
        if not service_key:
            return None
        return create_client(st.secrets["SUPABASE_URL"], service_key)
    except Exception:
        return None


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt(rounds=12)).decode("utf-8")


def valid_password(password: str) -> bool:
    return (
        len(password) >= 10
        and any(c.isupper() for c in password)
        and any(c.islower() for c in password)
        and any(c.isdigit() for c in password)
    )


def is_demo() -> bool:
    """Keep preview sessions completely isolated from Supabase."""
    user = st.session_state.get("user", {})
    return user.get("id") == "demo"


def fetch_table(name: str, fallback: pd.DataFrame) -> tuple[pd.DataFrame, bool]:
    if is_demo():
        return fallback.copy(), False
    client = db()
    if client:
        try:
            rows = client.table(name).select("*").execute().data
            return (pd.DataFrame(rows), True) if rows else (fallback.copy(), True)
        except Exception:
            pass
    return fallback.copy(), False


def upsert_rows(name: str, frame: pd.DataFrame) -> tuple[bool, str]:
    if is_demo():
        return False, "Demo mode is read-only. No Supabase data was changed."
    client = db()
    if not client:
        return False, "Supabase is unavailable. No data was changed."
    try:
        clean = frame.where(pd.notna(frame), None).to_dict("records")
        client.table(name).upsert(clean).execute()
        return True, f"{len(clean)} records saved."
    except Exception as exc:
        return False, f"Could not save records: {exc}"


def upload_file(bucket: str, path: str, data: bytes, content_type: str) -> tuple[bool, str]:
    if is_demo():
        return False, "Demo mode is read-only. No file was uploaded."
    client = db()
    if not client:
        return False, "Supabase storage is unavailable."
    try:
        client.storage.from_(bucket).upload(path, data, {"content-type": content_type, "upsert": "true"})
        return True, path
    except Exception as exc:
        return False, str(exc)


def extract_material_text(filename: str, data: bytes) -> str:
    extension = filename.lower().rsplit(".", 1)[-1]
    if extension == "pdf" and PdfReader:
        return "\n".join(page.extract_text() or "" for page in PdfReader(io.BytesIO(data)).pages)
    if extension == "docx" and Document:
        document = Document(io.BytesIO(data))
        return "\n".join(paragraph.text for paragraph in document.paragraphs)
    if extension == "pptx" and Presentation:
        presentation = Presentation(io.BytesIO(data))
        return "\n".join(
            shape.text
            for slide in presentation.slides
            for shape in slide.shapes
            if hasattr(shape, "text")
        )
    if extension in {"txt", "md", "csv"}:
        return data.decode("utf-8", errors="ignore")
    raise ValueError("Quiz generation supports PDF, DOCX, PPTX, TXT, MD and CSV materials.")


def generate_material_quiz(material: pd.Series, question_count: int) -> tuple[bool, str]:
    api_key = st.secrets.get("OPENAI_API_KEY")
    if not api_key or OpenAI is None:
        return False, "Add OPENAI_API_KEY to Streamlit secrets and install requirements to enable AI quiz generation."
    client = db()
    try:
        file_path = str(material["file_path"])
        data = client.storage.from_("materials").download(file_path)
        source_text = extract_material_text(file_path, data).strip()
        if len(source_text) < 200:
            return False, "The selected material does not contain enough readable text."
        source_text = source_text[:60_000]
        generated_questions: list[GeneratedQuestion] = []
        generated_title = f"{material['title']} quiz"
        openai_client = OpenAI(api_key=api_key)
        remaining = question_count
        batch_number = 1
        while remaining:
            batch_size = min(20, remaining)
            response = openai_client.responses.parse(
                model="gpt-5.6-luna",
                input=[
                    {
                        "role": "system",
                        "content": (
                            "Create a fair educational multiple-choice quiz using only the supplied "
                            "study material. Use four plausible options per question, one correct answer, "
                            "clear explanations, and no trick questions. Avoid repeating questions."
                        ),
                    },
                    {
                        "role": "user",
                        "content": (
                            f"Create exactly {batch_size} questions for batch {batch_number}. "
                            f"Do not repeat these existing questions: "
                            f"{[item.question for item in generated_questions]}\n\n"
                            f"Study material:\n{source_text}"
                        ),
                    },
                ],
                text_format=GeneratedQuiz,
            )
            generated = response.output_parsed
            if batch_number == 1 and generated.title:
                generated_title = generated.title
            generated_questions.extend(generated.questions)
            remaining -= batch_size
            batch_number += 1
        quiz_row = client.table("quizzes").insert({
            "material_id": int(material["id"]),
            "title": generated_title,
            "instructions": "Choose the best answer for each question.",
            "xp_reward": 25,
            "passing_score": 60,
            "status": "draft",
            "generated_by_ai": True,
            "created_by": st.session_state.user["id"],
        }).execute().data[0]
        questions = [
            {
                "quiz_id": quiz_row["id"],
                "position": index,
                "question": item.question,
                "options": item.options,
                "correct_index": item.correct_index,
                "explanation": item.explanation,
            }
            for index, item in enumerate(generated_questions, start=1)
        ]
        client.table("quiz_questions").insert(questions).execute()
        return True, f"Draft quiz created with {len(questions)} questions."
    except Exception as exc:
        return False, f"Quiz generation failed: {exc}"


def metric(label: str, value: str, note: str) -> None:
    st.markdown(
        f'<div class="metric-card"><div class="metric-label">{label}</div>'
        f'<div class="metric-value">{value}</div><div class="metric-note">{note}</div></div>',
        unsafe_allow_html=True,
    )


def polish_chart(fig: Any) -> Any:
    """Keep Plotly typography readable regardless of the active Streamlit theme."""
    fig.update_layout(
        font=dict(family="DM Sans, sans-serif", color="#242424", size=13),
        title_font=dict(family="Manrope, sans-serif", color="#171717"),
        legend=dict(font=dict(color="#242424")),
        hoverlabel=dict(
            bgcolor="#ffffff",
            bordercolor="#1f5daa",
            font=dict(family="DM Sans, sans-serif", color="#171717", size=13),
        ),
    )
    fig.update_xaxes(
        tickfont=dict(color="#3f3f3f"),
        title_font=dict(color="#3f3f3f"),
        linecolor="#cfcfcb",
        gridcolor="#e9e9e6",
        zerolinecolor="#cfcfcb",
    )
    fig.update_yaxes(
        tickfont=dict(color="#3f3f3f"),
        title_font=dict(color="#3f3f3f"),
        linecolor="#cfcfcb",
        gridcolor="#e9e9e6",
        zerolinecolor="#cfcfcb",
    )
    return fig


def heading(eyebrow: str, title: str, copy: str = "") -> None:
    st.markdown(
        f'<div class="hero"><div class="topline">{eyebrow}</div>'
        f"<h1>{title}</h1><p>{copy}</p></div>",
        unsafe_allow_html=True,
    )


def sidebar(role: str, name: str) -> str:
    with st.sidebar:
        if role == "Student":
            st.markdown('<span class="student-sidebar-marker"></span>', unsafe_allow_html=True)
        st.image(LOGO_IMAGE, width=112)
        st.caption("Experience-led learning")
        st.markdown(
            f'<div class="userbox"><b>{name}</b><br><small>{role} workspace</small></div>',
            unsafe_allow_html=True,
        )
        if is_demo():
            roles = ["Student", "Admin"]
            selected_role = st.selectbox(
                "Demo workspace",
                roles,
                index=roles.index(role),
                help="Preview XPLMS from a different role.",
            )
            if selected_role != role:
                st.session_state.user.update(
                    {"name": f"Demo {selected_role}", "role": selected_role}
                )
                st.rerun()
        elif st.session_state.user.get("role") == "Admin":
            preview_roles = ["Admin", "Student"]
            selected_preview = st.selectbox(
                "Preview workspace",
                preview_roles,
                index=preview_roles.index(role),
                help="Preview another workspace without changing your administrator account.",
            )
            if selected_preview != role:
                st.session_state.admin_preview_role = selected_preview
                st.rerun()
        menus = {
            "Student": [
                "Progress", "Materials", "Quiz", "Leaderboard", "Profile", "XP"
            ],
            "Admin": [
                "Student record",
                "Analysis background",
                "Analysis progress",
                "Analysis XP",
                "Analysis material",
                "CRUD",
                "Award XP",
                "User access",
                "Material",
                "Quiz",
            ],
        }
        labels = {
            item: item
            for items in menus.values()
            for item in items
        }
        chosen = st.radio("Navigation", menus[role], format_func=lambda x: labels[x], label_visibility="collapsed")
        st.divider()
        if st.button("Sign out", use_container_width=True):
            st.session_state.clear()
            st.rerun()
    return chosen


def login() -> None:
    left, centre, right = st.columns([1, 1.15, 1])
    with centre:
        st.markdown("<div style='height:6vh'></div>", unsafe_allow_html=True)
        st.image(LOGO_IMAGE, width=132)
        client = db()
        if not client:
            st.error("XPLMS server access is not configured. Add SUPABASE_SECRET_KEY to Streamlit secrets.")
            return
        try:
            existing_users = client.table("app_users").select("id", count="exact").limit(1).execute()
            user_count = existing_users.count or 0
        except Exception:
            st.error("The app_users table is not available. Run supabase_migration_002_app_users.sql first.")
            return

        if user_count == 0:
            st.info("Create the first XPLMS administrator. This setup appears only while no application users exist.")
            with st.form("bootstrap_admin"):
                full_name = st.text_input("Administrator name")
                email = st.text_input("Administrator email")
                password = st.text_input("New password", type="password")
                confirm = st.text_input("Confirm password", type="password")
                if st.form_submit_button("Create administrator", type="primary"):
                    if not full_name.strip() or "@" not in email:
                        st.error("Enter a valid name and email.")
                    elif password != confirm:
                        st.error("The passwords do not match.")
                    elif not valid_password(password):
                        st.error("Use at least 10 characters with uppercase, lowercase and a number.")
                    else:
                        try:
                            client.table("app_users").insert({
                                "email": email.strip().lower(),
                                "full_name": full_name.strip(),
                                "role": "admin",
                                "NO MATRIK": None,
                                "password_hash": hash_password(password),
                                "active": True,
                                "must_change_password": False,
                            }).execute()
                            st.success("Administrator created. Sign in below.")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Administrator could not be created: {exc}")
            return

        with st.container(border=True):
            email = st.text_input("Email address", placeholder="name@institution.edu")
            password = st.text_input("Password", type="password", placeholder="Your password")
            if st.button("Sign in", type="primary", use_container_width=True):
                try:
                    rows = (
                        client.table("app_users")
                        .select('id,email,full_name,role,"NO MATRIK",password_hash,active,must_change_password')
                        .ilike("email", email.strip())
                        .limit(1)
                        .execute()
                        .data
                    )
                    account = rows[0] if rows else None
                    if (
                        not account
                        or not account.get("active")
                        or not bcrypt.checkpw(
                            password.encode("utf-8"),
                            account["password_hash"].encode("utf-8"),
                        )
                    ):
                        raise ValueError("Invalid credentials")
                    st.session_state.user = {
                        "name": account["full_name"],
                        "email": account["email"],
                        "role": "Admin" if account["role"] == "lecturer" else account["role"].title(),
                        "id": str(account["id"]),
                        "no_matrik": account.get("NO MATRIK"),
                        "must_change_password": account.get("must_change_password", False),
                    }
                    client.table("app_users").update(
                        {"last_login_at": datetime.now().isoformat()}
                    ).eq("id", account["id"]).execute()
                    st.rerun()
                except Exception:
                    st.error("We couldn't sign you in. Check your email and password.")
            st.caption("Forgot your password? Contact your XPLMS administrator.")
        with st.expander("Preview the app"):
            st.caption("Demo workspaces use sample data and never change Supabase.")
            cols = st.columns(2)
            for col, role in zip(cols, ["Student", "Admin"]):
                if col.button(role, use_container_width=True):
                    st.session_state.user = {"name": f"Demo {role}", "email": "demo@xplms.edu", "role": role, "id": "demo"}
                    st.rerun()


def merged_students() -> tuple[pd.DataFrame, bool]:
    students, live_a = fetch_table("stud_background", DEMO_STUDENTS)
    progress, live_b = fetch_table("stud_progress", DEMO_PROGRESS)
    xp, live_c = fetch_table(
        "stud_xp", DEMO_PROGRESS[["student_id", "xp"]].copy()
    )
    key = next((x for x in ["NO MATRIK", "student_id", "stud_id", "id"] if x in students.columns and x in progress.columns), None)
    if not key:
        return students, live_a
    merged = students.merge(progress, on=key, how="left")
    xp_key = next((x for x in ["NO MATRIK", "student_id", "stud_id", "id"] if x in merged.columns and x in xp.columns), None)
    if xp_key:
        merged = merged.merge(xp, on=xp_key, how="left")
    if "XP" in merged.columns:
        merged["xp"] = pd.to_numeric(merged["XP"], errors="coerce").fillna(0)
        merged = merged.drop(columns=["XP"])
    return merged, live_a and live_b and live_c


def student_nickname(user: dict[str, Any]) -> str:
    """Return the student's preferred nickname with a safe account fallback."""
    background, _ = fetch_table("stud_background", DEMO_STUDENTS)
    matric = user.get("no_matrik")
    if matric and "NO MATRIK" in background.columns:
        matched = background[background["NO MATRIK"].astype(str) == str(matric)]
    elif is_demo():
        matched = background.head(1)
    else:
        matched = pd.DataFrame()
    if not matched.empty:
        row = matched.iloc[0]
        for column in ["NICKNAME PELAJAR", "nickname", "NAMA PELAJAR", "name"]:
            value = row.get(column)
            if pd.notna(value) and str(value).strip():
                return str(value).strip()
    return str(user.get("name", "Student"))


ASSESSMENT_COLUMNS = [
    "C1C2",
    "C5",
    "C8",
    "C9C10",
    "INDIVIDUAL ASSIGNMENT",
    "GROUP ASSIGNMENT",
    "UPS 1",
    "UPS 2",
    "UPS 3",
    # Demo/legacy compatibility; live XPLMS uses the assessment fields above.
    "assignment",
    "pb",
    "task",
]


def leaderboard_data(
    month: str | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, list[str], str]:
    """Build fair individual and class rankings from the XP ledger."""
    merged, _ = merged_students()
    events, _ = fetch_table("xp_events", pd.DataFrame())
    current_month = datetime.now(ZoneInfo("Asia/Kuala_Lumpur")).strftime("%Y-%m")
    available_months = [current_month]
    if not events.empty and "created_at" in events.columns:
        event_dates = pd.to_datetime(events["created_at"], errors="coerce", utc=True)
        event_months = (
            event_dates.dt.tz_convert("Asia/Kuala_Lumpur")
            .dt.strftime("%Y-%m")
            .dropna()
            .unique()
            .tolist()
        )
        available_months = sorted(set([*event_months, current_month]), reverse=True)
    selected_month = month if month in available_months else available_months[0]
    if merged.empty:
        return pd.DataFrame(), pd.DataFrame(), available_months, selected_month

    matric_col = next(
        (column for column in ["NO MATRIK", "student_id", "stud_id"] if column in merged.columns),
        None,
    )
    name_col = next(
        (
            column
            for column in [
                "NICKNAME PELAJAR", "nickname", "NAMA PELAJAR", "name", "student_name"
            ]
            if column in merged.columns
        ),
        None,
    )
    class_col = next(
        (column for column in ["KELAS", "CLASS", "class", "cohort"] if column in merged.columns),
        None,
    )
    if not matric_col:
        return pd.DataFrame(), pd.DataFrame(), available_months, selected_month

    individuals = pd.DataFrame({
        "NO MATRIK": merged[matric_col].astype(str),
        "Student": (
            merged[name_col].fillna("Unknown").astype(str)
            if name_col
            else merged[matric_col].astype(str)
        ),
        "Class": (
            merged[class_col].fillna("Unassigned").astype(str)
            if class_col
            else "Unassigned"
        ),
    })
    overall_values = (
        merged["xp"]
        if "xp" in merged.columns
        else pd.Series(0, index=merged.index)
    )
    individuals["Overall XP"] = pd.to_numeric(
        overall_values, errors="coerce"
    ).fillna(0)

    monthly_xp = pd.Series(dtype=float)
    if not events.empty and {"NO MATRIK", "points", "created_at"}.issubset(events.columns):
        event_dates = pd.to_datetime(events["created_at"], errors="coerce", utc=True)
        event_month = event_dates.dt.tz_convert("Asia/Kuala_Lumpur").dt.strftime("%Y-%m")
        month_events = events[event_month == selected_month].copy()
        month_events["points"] = pd.to_numeric(month_events["points"], errors="coerce").fillna(0)
        monthly_xp = month_events.groupby("NO MATRIK")["points"].sum()
    individuals["Monthly XP"] = (
        individuals["NO MATRIK"].map(monthly_xp).fillna(0)
    )

    assessment_columns = [
        column for column in ASSESSMENT_COLUMNS if column in merged.columns
    ]
    if assessment_columns:
        numeric_marks = merged[assessment_columns].apply(
            pd.to_numeric, errors="coerce"
        )
        individuals["Progress Average"] = numeric_marks.mean(axis=1).round(2)
    else:
        individuals["Progress Average"] = pd.NA

    individuals["Monthly Rank"] = (
        individuals["Monthly XP"].rank(method="min", ascending=False).astype(int)
    )
    individuals["Overall Rank"] = (
        individuals["Overall XP"].rank(method="min", ascending=False).astype(int)
    )
    progress_rank = individuals["Progress Average"].rank(method="min", ascending=False)
    individuals["Progress Rank"] = progress_rank.astype("Int64")
    individuals["Badge"] = pd.cut(
        individuals["Overall XP"],
        bins=[-1, 99, 299, 599, 999, float("inf")],
        labels=["None", "Starter", "Active learner", "Achiever", "XP Elite"],
    ).astype(str)

    classes = (
        individuals.groupby("Class", dropna=False)
        .agg(
            Students=("NO MATRIK", "nunique"),
            **{
                "Monthly XP Total": ("Monthly XP", "sum"),
                "Monthly XP Average": ("Monthly XP", "mean"),
                "Overall XP Total": ("Overall XP", "sum"),
                "Overall XP Average": ("Overall XP", "mean"),
                "Progress Average": ("Progress Average", "mean"),
            },
        )
        .reset_index()
    )
    classes["Monthly Rank"] = (
        classes["Monthly XP Average"].rank(method="min", ascending=False).astype(int)
    )
    classes["Overall Rank"] = (
        classes["Overall XP Average"].rank(method="min", ascending=False).astype(int)
    )
    classes["Progress Rank"] = (
        classes["Progress Average"].rank(method="min", ascending=False).astype("Int64")
    )
    classes["Class Badge"] = pd.cut(
        classes["Overall XP Average"],
        bins=[-1, 99, 299, 599, 999, float("inf")],
        labels=["None", "Starter", "Active learner", "Achiever", "XP Elite"],
    ).astype(str)
    return individuals, classes, available_months, selected_month


def progress_page() -> None:
    heading("Academic record", "Progress & results", "Review your latest marks and learning progress.")
    progress, _ = fetch_table("stud_progress", DEMO_PROGRESS)
    user = st.session_state.get("user", {})
    if user.get("no_matrik") and "NO MATRIK" in progress.columns:
        own = progress[progress["NO MATRIK"] == user["no_matrik"]]
        row = own.iloc[0] if not own.empty else None
    else:
        row = progress.iloc[0] if not progress.empty else None
    if row is None:
        st.info("No progress record is linked to this account.")
        return
    score_columns = [
        c for c in [
            "C1C2", "C5", "C8", "C9C10",
            "INDIVIDUAL ASSIGNMENT", "GROUP ASSIGNMENT",
            "UPS 1", "UPS 2", "UPS 3",
            "assignment", "pb", "task",
        ]
        if c in row.index
    ]
    results = []
    for column in score_columns:
        value = pd.to_numeric(pd.Series([row[column]]), errors="coerce").iloc[0]
        zone_column = f"{column}_ZONE"
        mark = None if pd.isna(value) else float(value)
        results.append({
            "Component": column.replace("_", " "),
            "Mark": mark,
            "Zone": row.get(zone_column, "—") if not pd.isna(row.get(zone_column, None)) else "—",
            "Status": "Available" if mark is not None else "Pending",
        })
    completed = [item["Mark"] for item in results if item["Mark"] is not None]
    a, b, c = st.columns(3)
    with a: metric("Results available", str(len(completed)), f"of {len(results)} components")
    with b: metric("Current average", f"{sum(completed)/len(completed):.1f}" if completed else "—", "Based on available marks")
    with c: metric("Matric number", str(row.get("NO MATRIK", user.get("no_matrik", "—"))), "Student identifier")

    st.subheader("Assessment results")
    if not completed:
        st.info("Marks have not been published yet.")
    results_table = pd.DataFrame(results)
    results_table["Mark"] = results_table["Mark"].apply(
        lambda value: "—" if pd.isna(value) else f"{float(value):.1f}"
    )
    st.dataframe(
        results_table,
        hide_index=True,
        width="stretch",
        column_config={
            "Component": st.column_config.TextColumn("Assessment", width="large"),
            "Mark": st.column_config.TextColumn("Mark", width="small"),
            "Zone": st.column_config.TextColumn("Zone", width="medium"),
            "Status": st.column_config.TextColumn("Status", width="small"),
        },
    )


def profile_page() -> None:
    heading("Student information", "My profile", "View the background information linked to your matric number.")
    background, _ = fetch_table("stud_background", DEMO_STUDENTS)
    user = st.session_state.get("user", {})
    if user.get("no_matrik") and "NO MATRIK" in background.columns:
        own = background[background["NO MATRIK"] == user["no_matrik"]]
        row = own.iloc[0] if not own.empty else None
    else:
        row = background.iloc[0] if not background.empty else None
    if row is None:
        st.info("No student background record is linked to this account.")
        return

    name = row.get(
        "NICKNAME PELAJAR",
        row.get("nickname", row.get("NAMA PELAJAR", row.get("name", user.get("name", "—")))),
    )
    matric = row.get("NO MATRIK", row.get("student_id", user.get("no_matrik", "—")))
    student_class = row.get("KELAS", row.get("cohort", "—"))
    system = row.get("SISTEM", row.get("programme", "—"))
    a, b, c, d = st.columns(4)
    with a: metric("Student name", str(name), "Registered name")
    with b: metric("No Matrik", str(matric), "Student identifier")
    with c: metric("Class", str(student_class), "Current class")
    with d: metric("System", str(system), "Academic system")

    hidden = {"id", "created_at", "updated_at"}
    details = [
        {
            "Field": str(column).replace("_", " ").title(),
            "Information": "—" if pd.isna(value) else str(value),
        }
        for column, value in row.items()
        if column not in hidden
    ]
    st.subheader("Student background")
    st.dataframe(pd.DataFrame(details), hide_index=True, width="stretch")


def xp_badge_page() -> None:
    heading(
        "Recognition request",
        "XP",
        "Request XP for a verified learning activity.",
    )
    user = st.session_state.get("user", {})
    st.write(
        "Submit evidence for consultation, class participation, commitment or "
        "a study group. An Admin will review the request."
    )
    with st.form("xp_proof_claim", clear_on_submit=True):
        claim_type = st.selectbox(
            "XP type",
            ["consultation", "class_participation", "commitment", "study_group"],
            format_func=lambda value: value.replace("_", " ").title(),
        )
        title = st.text_input("Request title")
        description = st.text_area("Describe the activity and what you learned")
        proof = st.file_uploader("Proof image", type=["jpg", "jpeg", "png", "webp"])
        if st.form_submit_button("Send for approval", type="primary"):
            if not title.strip() or not description.strip() or not proof:
                st.error("Complete the title, description and proof image.")
            elif is_demo():
                st.success("Demo request submitted for Admin approval.")
            else:
                client = db()
                path = f"{user['id']}/{datetime.now().timestamp()}-{proof.name}"
                ok, result = upload_file("xp-proofs", path, proof.getvalue(), proof.type)
                if not ok:
                    st.error(result)
                else:
                    try:
                        client.table("xp_claims").insert({
                            "student_user_id": user["id"],
                            "NO MATRIK": user["no_matrik"],
                            "claim_type": claim_type,
                            "title": title.strip(),
                            "description": description.strip(),
                            "proof_path": path,
                            "status": "pending",
                        }).execute()
                        st.success("Request submitted. An Admin will review it.")
                    except Exception as exc:
                        st.error(f"Request could not be submitted: {exc}")

    claims, live = fetch_table("xp_claims", pd.DataFrame())
    if live and "NO MATRIK" in claims.columns:
        claims = claims[claims["NO MATRIK"].astype(str) == str(user.get("no_matrik"))]
    st.subheader("My XP requests")
    if claims.empty:
        st.info("No XP requests submitted yet.")
    else:
        visible = [
            column for column in
            ["created_at", "claim_type", "title", "status", "admin_note"]
            if column in claims.columns
        ]
        st.dataframe(
            claims[visible].sort_values("created_at", ascending=False),
            hide_index=True,
            width="stretch",
        )


def student_leaderboard_page() -> None:
    heading(
        "Healthy competition",
        "Leaderboard",
        "Compare monthly and overall XP, plus academic progress.",
    )
    _, _, months, default_month = leaderboard_data()
    selected_month = st.selectbox(
        "XP month",
        months,
        index=months.index(default_month),
    )
    individuals, classes, _, _ = leaderboard_data(selected_month)
    if individuals.empty:
        st.info("Leaderboard data is not available.")
        return
    user = st.session_state.user
    own_matric = str(user.get("no_matrik", ""))
    own = individuals[individuals["NO MATRIK"].astype(str) == own_matric]
    if own.empty and is_demo():
        own = individuals.head(1)
    if not own.empty:
        row = own.iloc[0]
        badge_thresholds = [
            (100, "Starter"),
            (300, "Active learner"),
            (600, "Achiever"),
            (1000, "XP Elite"),
        ]
        captured = [
            badge for threshold, badge in badge_thresholds
            if float(row["Overall XP"]) >= threshold
        ]
        a, b, c, d, e = st.columns(5)
        with a: metric("Monthly XP", f"{int(row['Monthly XP']):,}", selected_month)
        with b: metric("Overall XP", f"{int(row['Overall XP']):,}", "Cumulative")
        with c: metric("Badges captured", str(len(captured)), row["Badge"])
        with d: metric("Monthly rank", f"#{int(row['Monthly Rank'])}", selected_month)
        with e:
            progress_rank = row["Progress Rank"]
            metric(
                "Progress rank",
                f"#{int(progress_rank)}" if pd.notna(progress_rank) else "—",
                f"{row['Progress Average']:.1f}" if pd.notna(row["Progress Average"]) else "No marks",
            )
        st.subheader("Captured badges")
        if captured:
            badge_columns = st.columns(len(captured))
            for column, badge in zip(badge_columns, captured):
                with column:
                    st.markdown(
                        f'<div class="metric-card" style="text-align:center">'
                        f'<div class="metric-value">✓</div><b>{badge}</b>'
                        f'<div class="metric-note">Captured</div></div>',
                        unsafe_allow_html=True,
                    )
        else:
            st.info("No badges captured yet. The first badge unlocks at 100 overall XP.")

    xp_tab, progress_tab = st.tabs(["XP leaderboard", "Progress leaderboard"])
    with xp_tab:
        st.subheader(f"Individual · {selected_month}")
        monthly = individuals.sort_values(
            ["Monthly Rank", "Overall Rank", "Student"]
        )[
            [
                "Monthly Rank", "Student", "NO MATRIK", "Class",
                "Monthly XP", "Overall XP", "Badge",
            ]
        ]
        st.dataframe(monthly, hide_index=True, width="stretch")
        st.subheader("Individual · Overall")
        overall = individuals.sort_values(["Overall Rank", "Student"])[
            [
                "Overall Rank", "Student", "NO MATRIK", "Class",
                "Overall XP", "Monthly XP", "Badge",
            ]
        ]
        st.dataframe(overall, hide_index=True, width="stretch")
        st.subheader(f"Class · {selected_month}")
        class_monthly = classes.sort_values(["Monthly Rank", "Class"])[
            [
                "Monthly Rank", "Class", "Students", "Monthly XP Average",
                "Monthly XP Total", "Overall XP Average", "Class Badge",
            ]
        ]
        st.dataframe(class_monthly, hide_index=True, width="stretch")
        st.subheader("Class · Overall")
        class_overall = classes.sort_values(["Overall Rank", "Class"])[
            [
                "Overall Rank", "Class", "Students", "Overall XP Average",
                "Overall XP Total", "Monthly XP Average", "Class Badge",
            ]
        ]
        st.dataframe(class_overall, hide_index=True, width="stretch")
    with progress_tab:
        st.subheader("Individual progress")
        progress = individuals.dropna(subset=["Progress Average"]).sort_values(
            ["Progress Rank", "Student"]
        )[
            ["Progress Rank", "Student", "NO MATRIK", "Class", "Progress Average"]
        ]
        st.dataframe(progress, hide_index=True, width="stretch")
        st.subheader("Class progress")
        class_progress = classes.dropna(subset=["Progress Average"]).sort_values(
            ["Progress Rank", "Class"]
        )[
            ["Progress Rank", "Class", "Students", "Progress Average"]
        ]
        st.dataframe(class_progress, hide_index=True, width="stretch")


def materials_page(lecturer: bool = False, quiz_tools: bool = False) -> None:
    heading(
        "Knowledge library",
        "Study materials",
        "Browse learning resources by chapter and material type.",
    )
    materials, live = fetch_table("materials", EMPTY_MATERIALS)
    chapters = [1, 2, 5, 8, 9, 10]
    material_types = ["Infographic", "Notes", "Exercise", "Extra", "Reference", "Other"]

    if lecturer:
        with st.expander("Upload new material"):
            with st.form("material_upload"):
                title = st.text_input("Resource title")
                course = st.text_input("Course or subject", value="Mathematics")
                chapter = st.segmented_control(
                    "Chapter",
                    chapters,
                    default=chapters[0],
                    format_func=lambda value: f"Chapter {value}",
                )
                material_type = st.segmented_control(
                    "Material type",
                    material_types,
                    default=material_types[0],
                )
                file = st.file_uploader("File", type=["pdf", "docx", "pptx", "xlsx", "txt", "md", "csv", "zip"])
                if st.form_submit_button("Share with students", type="primary"):
                    if not title.strip():
                        st.error("Enter a resource title.")
                    elif not file:
                        st.error("Choose a file.")
                    else:
                        client = db()
                        try:
                            client.table("materials").select("chapter,material_type").limit(1).execute()
                        except Exception:
                            st.error("Run supabase_migration_005_material_classification.sql before uploading materials.")
                        else:
                            storage_path = (
                                f"chapter-{chapter}/{material_type.lower()}/"
                                f"{datetime.now().timestamp()}-{file.name}"
                            )
                            ok, path = upload_file(
                                "materials", storage_path, file.getvalue(), file.type
                            )
                            if ok:
                                saved, message = upsert_rows(
                                    "materials",
                                    pd.DataFrame([{
                                        "title": title.strip(),
                                        "course": course.strip() or "Mathematics",
                                        "chapter": int(chapter),
                                        "material_type": material_type,
                                        "file_path": path,
                                        "type": file.name.rsplit(".", 1)[-1].upper(),
                                        "lecturer_id": st.session_state.user["id"],
                                    }]),
                                )
                                if saved:
                                    st.success("Material shared with students.")
                                    st.rerun()
                                else:
                                    st.error(message)
                            else:
                                st.error(path)

    filter_a, filter_b = st.columns(2)
    with filter_a:
        chapter_filter = st.segmented_control(
            "Chapter",
            ["All", *chapters],
            default="All",
            format_func=lambda value: (
                "All chapters" if value == "All" else f"Chapter {value}"
            ),
            key=f"{'admin' if lecturer else 'student'}_material_chapter",
        )
    with filter_b:
        type_filter = st.selectbox(
            "Material type",
            ["All types", *material_types],
            key=f"{'admin' if lecturer else 'student'}_material_type",
        )

    query = st.text_input("Search materials", placeholder="Search by title, subject or type…")
    if query:
        materials = materials[materials.astype(str).apply(lambda row: row.str.contains(query, case=False).any(), axis=1)]
    if chapter_filter != "All" and "chapter" in materials.columns:
        chapter_values = pd.to_numeric(materials["chapter"], errors="coerce")
        materials = materials[chapter_values == int(chapter_filter)]
    if type_filter != "All types" and "material_type" in materials.columns:
        materials = materials[materials["material_type"] == type_filter]

    if materials.empty:
        st.info(
            "No materials have been uploaded yet."
            if not query and chapter_filter == "All" and type_filter == "All types"
            else "No materials match the selected filters."
        )

    for i, item in materials.iterrows():
        with st.container(border=True):
            c1, c2, c3, c4 = st.columns([5, 2.4, 1.2, 1.6])
            c1.markdown(f"**{item['title']}**  \n{item['course']}")
            c2.caption(
                f"Chapter {item.get('chapter', '—')} · "
                f"{item.get('material_type', 'Other')} · {item.get('type', 'FILE')}"
            )
            if live and item.get("file_path"):
                try:
                    file_data = db().storage.from_("materials").download(str(item["file_path"]))
                    c3.download_button(
                        "Download",
                        data=file_data,
                        file_name=str(item["file_path"]).rsplit("/", 1)[-1],
                        key=f"material_{i}",
                        width="stretch",
                    )
                except Exception:
                    c3.button(
                        "Unavailable",
                        key=f"material_{i}",
                        disabled=True,
                        width="stretch",
                    )
            if quiz_tools and live and item.get("file_path") and c4.button("Generate quiz", key=f"quiz_material_{i}"):
                with st.spinner("Generating a draft quiz from this material…"):
                    ok, message = generate_material_quiz(item, 10)
                (st.success if ok else st.error)(message)

    if quiz_tools:
        st.subheader("Quiz management")
        quizzes, quiz_live = fetch_table("quizzes", pd.DataFrame())
        if quizzes.empty:
            st.info("No quizzes created yet.")
        else:
            for _, quiz in quizzes.iterrows():
                with st.container(border=True):
                    q1, q2, q3 = st.columns([5, 2, 1.5])
                    q1.markdown(f"**{quiz['title']}**")
                    q2.caption(f"{quiz['status'].title()} · {quiz['xp_reward']} XP")
                    if quiz["status"] == "draft" and q3.button("Publish", key=f"publish_quiz_{quiz['id']}"):
                        db().table("quizzes").update({"status": "published"}).eq("id", quiz["id"]).execute()
                        st.success("Quiz published to students.")
                        st.rerun()


def admin_quiz_page() -> None:
    heading(
        "Assessment builder",
        "AI quiz generation",
        "Generate up to 100 multiple-choice questions from an uploaded material.",
    )
    materials, live = fetch_table("materials", EMPTY_MATERIALS)
    if materials.empty:
        st.info("Upload at least one material before generating a quiz.")
    else:
        labels = {
            int(row["id"]): (
                f"Chapter {row.get('chapter', '—')} · "
                f"{row.get('material_type', 'Other')} · {row['title']}"
            )
            for _, row in materials.iterrows()
        }
        with st.form("admin_generate_quiz"):
            material_id = st.selectbox(
                "Source material",
                list(labels),
                format_func=lambda value: labels[value],
            )
            question_count = st.number_input(
                "Number of questions",
                min_value=1,
                max_value=100,
                value=20,
                step=5,
            )
            generate = st.form_submit_button("Generate draft quiz", type="primary")
            if generate:
                selected = materials[materials["id"] == material_id].iloc[0]
                with st.spinner(
                    f"Generating {int(question_count)} questions in batches…"
                ):
                    ok, message = generate_material_quiz(
                        selected, int(question_count)
                    )
                (st.success if ok else st.error)(message)

    st.subheader("Quiz library")
    quizzes, _ = fetch_table("quizzes", pd.DataFrame())
    if quizzes.empty:
        st.info("No quizzes created yet.")
        return
    for _, quiz in quizzes.iterrows():
        with st.container(border=True):
            left, middle, right = st.columns([5, 2, 1.5])
            left.markdown(f"**{quiz['title']}**")
            middle.caption(
                f"{str(quiz['status']).title()} · {quiz.get('xp_reward', 0)} XP"
            )
            if quiz["status"] == "draft" and right.button(
                "Publish", key=f"admin_publish_quiz_{quiz['id']}"
            ):
                db().table("quizzes").update({"status": "published"}).eq(
                    "id", quiz["id"]
                ).execute()
                st.success("Quiz published to students.")
                st.rerun()


def quiz_page() -> None:
    heading(
        "Daily knowledge check",
        "Chapter quiz",
        "Answer 10 random questions per chapter each day and earn automatic XP.",
    )
    user = st.session_state.user
    today = datetime.now(ZoneInfo("Asia/Kuala_Lumpur")).date().isoformat()
    materials, materials_live = fetch_table("materials", EMPTY_MATERIALS)
    quizzes, quizzes_live = fetch_table("quizzes", pd.DataFrame())
    live = materials_live and quizzes_live

    if live:
        quizzes = quizzes[quizzes.get("status") == "published"].copy()
        if quizzes.empty or materials.empty:
            st.info("No published chapter question banks are available.")
            return
        quiz_materials = quizzes.merge(
            materials[["id", "chapter"]],
            left_on="material_id",
            right_on="id",
            how="left",
            suffixes=("_quiz", "_material"),
        )
        quiz_materials = quiz_materials.dropna(subset=["chapter"])
        chapters = sorted(quiz_materials["chapter"].astype(int).unique().tolist())
    else:
        chapters = [1]
        quiz_materials = pd.DataFrame(
            [{"id_quiz": 1, "chapter": 1, "title": "Daily Chapter 1 quiz"}]
        )

    if not chapters:
        st.info("Published quizzes must be linked to materials with a chapter.")
        return
    chapter = st.selectbox(
        "Chapter",
        chapters,
        format_func=lambda value: f"Chapter {value}",
    )

    if live:
        client = db()
        try:
            previous = (
                client.table("quiz_attempts")
                .select("score,correct_count,total_questions,xp_awarded,completed_at")
                .eq("student_user_id", user["id"])
                .eq("chapter", int(chapter))
                .eq("attempt_date", today)
                .limit(1)
                .execute()
                .data
            )
        except Exception:
            st.error(
                "Run supabase_migration_006_revised_xp_daily_quiz.sql "
                "before using daily quizzes."
            )
            return
        if previous:
            result = previous[0]
            st.success(
                f"Today's Chapter {chapter} quiz is complete · "
                f"{int(result['correct_count'])}/{int(result['total_questions'])} correct · "
                f"{int(result['xp_awarded'])} XP awarded."
            )
            return
        chapter_quizzes = quiz_materials[
            quiz_materials["chapter"].astype(int) == int(chapter)
        ]
        quiz_ids = chapter_quizzes["id_quiz"].astype(int).tolist()
        questions = []
        for quiz_id in quiz_ids:
            rows = (
                client.table("quiz_questions")
                .select("*")
                .eq("quiz_id", quiz_id)
                .execute()
                .data
            )
            questions.extend(rows)
    else:
        quiz_ids = [1]
        questions = [
            {
                "id": index,
                "quiz_id": 1,
                "question": f"Demo question {index}: what is {index} + 1?",
                "options": [str(index), str(index + 1), str(index + 2), str(index + 3)],
                "correct_index": 1,
                "explanation": f"{index} + 1 equals {index + 1}.",
            }
            for index in range(1, 11)
        ]

    if len(questions) < 10:
        st.info(
            f"Chapter {chapter} currently has {len(questions)} questions. "
            "An Admin must generate at least 10 before the daily quiz opens."
        )
        return

    seed_text = f"{user['id']}:{chapter}:{today}"
    seed = int(hashlib.sha256(seed_text.encode("utf-8")).hexdigest()[:8], 16)
    daily_questions = (
        pd.DataFrame(questions).sample(n=10, random_state=seed).to_dict("records")
    )
    st.caption(
        "XP rule: 1 XP for an attempted question; a correct answer earns 2 XP instead."
    )
    with st.form(f"daily_quiz_{chapter}_{today}"):
        answers = {}
        for number, question in enumerate(daily_questions, start=1):
            options = question["options"]
            if isinstance(options, str):
                options = json.loads(options)
            answers[str(question["id"])] = st.radio(
                f"{number}. {question['question']}",
                range(len(options)),
                format_func=lambda index, opts=options: opts[index],
                index=None,
                key=f"daily_answer_{chapter}_{today}_{question['id']}",
            )
        if st.form_submit_button("Submit today's quiz", type="primary"):
            if any(answer is None for answer in answers.values()):
                st.error("Answer all 10 questions before submitting.")
            else:
                correct = sum(
                    int(answers[str(question["id"])])
                    == int(question["correct_index"])
                    for question in daily_questions
                )
                total_questions = len(daily_questions)
                score = correct / total_questions * 100
                xp_awarded = total_questions + correct
                if not live:
                    st.success(
                        f"Demo result: {correct}/10 correct · {xp_awarded} XP."
                    )
                else:
                    try:
                        attempt = client.table("quiz_attempts").insert({
                            "quiz_id": int(quiz_ids[0]),
                            "student_user_id": user["id"],
                            "NO MATRIK": user["no_matrik"],
                            "answers": answers,
                            "score": score,
                            "correct_count": correct,
                            "total_questions": total_questions,
                            "passed": score >= 60,
                            "chapter": int(chapter),
                            "attempt_date": today,
                            "xp_awarded": xp_awarded,
                        }).execute().data[0]
                        event = client.table("xp_events").insert({
                            "NO MATRIK": user["no_matrik"],
                            "rule_code": "quiz_completion",
                            "points": xp_awarded,
                            "source_id": f"daily-quiz-{chapter}-{today}",
                            "reason": (
                                f"Daily Chapter {chapter} quiz: "
                                f"{correct}/{total_questions} correct"
                            ),
                            "award_mode": "automatic",
                            "awarded_by": None,
                        }).execute().data[0]
                        client.table("quiz_attempts").update({
                            "xp_event_id": event["id"]
                        }).eq("id", attempt["id"]).execute()
                        st.success(
                            f"{correct}/10 correct · {xp_awarded} XP awarded."
                        )
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Quiz result could not be saved: {exc}")


def award_xp_page() -> None:
    heading(
        "Recognition",
        "Award experience points",
        "Award consultation, class participation and commitment, or review student requests.",
    )
    students, students_live = merged_students()
    rules_fallback = pd.DataFrame(
        [
            {"code": "consultation", "name": "Consultation", "default_points": 20, "award_mode": "manual"},
            {"code": "class_participation", "name": "Class participation", "default_points": 10, "award_mode": "manual"},
            {"code": "commitment", "name": "Commitment", "default_points": 10, "award_mode": "manual"},
        ]
    )
    rules, rules_live = fetch_table("xp_rules", rules_fallback)
    if "award_mode" in rules.columns:
        rules = rules[rules["award_mode"] == "manual"]
    if "code" in rules.columns:
        rules = rules[
            rules["code"].isin(
                ["consultation", "class_participation", "commitment"]
            )
        ]
    if "active" in rules.columns:
        rules = rules[rules["active"] == True]  # noqa: E712

    matric_col = next((c for c in ["NO MATRIK", "student_id"] if c in students.columns), None)
    name_col = next((c for c in ["NAMA PELAJAR", "name", "student_name"] if c in students.columns), None)
    if not matric_col or students.empty or rules.empty:
        st.info("Student profiles and active manual XP rules are required before XP can be awarded.")
        return

    choices = students[[matric_col] + ([name_col] if name_col else [])].drop_duplicates()
    labels = {
        str(row[matric_col]): (
            f"{row[name_col]} · {row[matric_col]}" if name_col else str(row[matric_col])
        )
        for _, row in choices.iterrows()
    }
    rule_lookup = {str(row["code"]): row for _, row in rules.iterrows()}

    with st.form("manual_xp_award", clear_on_submit=True):
        matric = st.selectbox("Student", list(labels), format_func=lambda value: labels[value])
        rule_code = st.selectbox(
            "Award category",
            list(rule_lookup),
            format_func=lambda value: str(rule_lookup[value].get("name", value)),
        )
        default_points = int(rule_lookup[rule_code].get("default_points", 10))
        points = st.number_input("XP points", min_value=1, max_value=1000, value=default_points)
        reason = st.text_area(
            "Reason",
            placeholder="Give a concise, specific reason for this recognition.",
        )
        confirmed = st.checkbox("I confirm this award is accurate and appropriate.")
        submitted = st.form_submit_button(
            "Award XP", type="primary", disabled=not confirmed
        )
        if submitted:
            if not reason.strip():
                st.error("Add a reason before awarding XP.")
            elif is_demo():
                st.success(f"Demo award recorded: +{points} XP for {labels[matric]}.")
            else:
                client = db()
                try:
                    if st.session_state.user.get("role") != "Admin":
                        raise PermissionError("Only administrators can award XP.")
                    client.table("xp_events").insert({
                        "NO MATRIK": matric,
                        "rule_code": rule_code,
                        "points": int(points),
                        "source_id": f"manual-{datetime.now().timestamp()}",
                        "reason": reason.strip(),
                        "award_mode": "manual",
                        "awarded_by": st.session_state.user["id"],
                    }).execute()
                    st.success(f"{points} XP awarded to {labels[matric]}.")
                except Exception as exc:
                    st.error(f"XP could not be awarded: {exc}")

    st.subheader("Student XP requests")
    claims_fallback = pd.DataFrame(
        [{
            "id": 1, "NO MATRIK": "S24001", "claim_type": "study_group",
            "title": "Algebra study group",
            "description": "Discussed simultaneous equations with classmates.",
            "proof_path": "", "status": "pending", "created_at": "2026-07-29",
        }]
    )
    claims, claims_live = fetch_table("xp_claims", claims_fallback)
    if "status" in claims.columns:
        claims = claims[claims["status"] == "pending"]
    if claims.empty:
        st.info("No pending XP claims.")
    else:
        client = db()
        for _, claim in claims.iterrows():
            claim_label = str(claim.get("claim_type", "study_group")).replace("_", " ").title()
            with st.expander(
                f"{claim_label} · {claim.get('title', 'XP request')} · "
                f"{claim.get('NO MATRIK', '')}"
            ):
                st.write(claim.get("description", ""))
                st.caption(f"Submitted {claim.get('created_at', 'recently')}")
                proof_path = claim.get("proof_path")
                if claims_live and proof_path:
                    try:
                        signed = client.storage.from_("xp-proofs").create_signed_url(str(proof_path), 600)
                        url = signed.get("signedURL") or signed.get("signedUrl")
                        if url:
                            st.image(url, caption="Submitted proof", width=420)
                    except Exception:
                        st.warning("The proof image could not be displayed.")
                admin_note = st.text_input("Admin note", key=f"claim_note_{claim['id']}")
                approve_col, reject_col = st.columns(2)
                if approve_col.button("Approve and award XP", type="primary", key=f"approve_claim_{claim['id']}"):
                    if not claims_live:
                        st.success("Demo claim approved.")
                    else:
                        try:
                            rule_code = str(claim.get("claim_type", "study_group"))
                            rule = client.table("xp_rules").select("default_points").eq("code", rule_code).limit(1).execute().data
                            points = int(rule[0]["default_points"]) if rule else 20
                            event = client.table("xp_events").insert({
                                "NO MATRIK": claim["NO MATRIK"],
                                "rule_code": rule_code,
                                "points": points,
                                "source_id": f"claim-{claim['id']}",
                                "reason": claim["title"],
                                "award_mode": "manual",
                                "awarded_by": st.session_state.user["id"],
                            }).execute().data[0]
                            client.table("xp_claims").update({
                                "status": "approved",
                                "admin_note": admin_note,
                                "reviewed_by": st.session_state.user["id"],
                                "reviewed_at": datetime.now().isoformat(),
                                "xp_event_id": event["id"],
                            }).eq("id", claim["id"]).execute()
                            st.success(f"Claim approved and {points} XP awarded.")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Claim approval failed: {exc}")
                if reject_col.button("Reject claim", key=f"reject_claim_{claim['id']}"):
                    if not claims_live:
                        st.info("Demo claim rejected.")
                    else:
                        client.table("xp_claims").update({
                            "status": "rejected",
                            "admin_note": admin_note,
                            "reviewed_by": st.session_state.user["id"],
                            "reviewed_at": datetime.now().isoformat(),
                        }).eq("id", claim["id"]).execute()
                        st.info("Claim rejected.")
                        st.rerun()

    st.subheader("How XP is awarded")
    st.markdown(
        "- **Admin awards:** consultation, class participation and commitment.\n"
        "- **Admin-approved requests:** consultation, class participation, commitment or study group with proof.\n"
        "- **Automatic awards:** daily in-app chapter quizzes."
    )
    if not (students_live and rules_live):
        st.caption("Showing the XP award workflow with demo data.")


def students_page(admin: bool = False) -> None:
    merged, live = merged_students()
    heading("Records", "Student records" if admin else "Students", "Search, review and maintain trusted learner information.")
    query = st.text_input("Search records", placeholder="Name, ID, email or programme…")
    shown = merged
    if query:
        shown = merged[merged.astype(str).apply(lambda row: row.str.contains(query, case=False).any(), axis=1)]
    st.dataframe(shown, use_container_width=True, hide_index=True)
    if admin:
        st.subheader("Edit records")
        edited = st.data_editor(merged, use_container_width=True, num_rows="dynamic", hide_index=True, key="student_editor")
        if st.button("Save changes", type="primary"):
            ok, message = upsert_rows("stud_background", edited[[c for c in edited.columns if c in DEMO_STUDENTS.columns]])
            (st.success if ok else st.warning)(message)
    if not live: st.caption("Sample data is shown because the live tables are unavailable.")


def admin_student_records_page() -> None:
    heading(
        "Student information",
        "Student record",
        "Combined background, progress and XP records with field-level filters.",
    )
    merged, live = merged_students()
    technical_columns = [
        column
        for column in merged.columns
        if (
            str(column).lower() in {"id", "created_at", "updated_at"}
            or str(column).lower().startswith(("id_", "created_at_", "updated_at_"))
        )
    ]
    merged = merged.drop(columns=technical_columns, errors="ignore")
    if merged.empty:
        st.info("No student records are available.")
        return
    search = st.text_input("Search records", placeholder="Name, No Matrik or any value…")
    categorical = [
        column
        for column in merged.columns
        if merged[column].nunique(dropna=True) <= 30
        and column not in {"id", "created_at", "updated_at"}
    ]
    slicer_fields = st.multiselect(
        "Dropdown slicers",
        categorical,
        max_selections=4,
        placeholder="Select fields such as class, programme, system or zone",
    )
    filtered = merged.copy()
    slicer_columns = st.columns(max(1, min(4, len(slicer_fields))))
    for index, field in enumerate(slicer_fields):
        values = sorted(filtered[field].dropna().astype(str).unique().tolist())
        selected = slicer_columns[index % len(slicer_columns)].selectbox(
            field, ["All", *values], key=f"record_slicer_{field}"
        )
        if selected != "All":
            filtered = filtered[filtered[field].astype(str) == selected]
    if search:
        filtered = filtered[
            filtered.astype(str).apply(
                lambda row: row.str.contains(search, case=False, na=False).any(),
                axis=1,
            )
        ]
    a, b = st.columns(2)
    with a:
        metric("Students shown", f"{len(filtered):,}", f"of {len(merged):,} records")
    with b:
        metric("Available fields", f"{len(merged.columns):,}", "Background, progress and XP")
    st.dataframe(filtered, hide_index=True, width="stretch")
    if not live:
        st.caption("Sample data is shown because live tables are unavailable.")


def analysis_background_page() -> None:
    heading(
        "Student intelligence",
        "Analysis background",
        "Understand the composition of the student cohort.",
    )
    background, _ = fetch_table("stud_background", DEMO_STUDENTS)
    if background.empty:
        st.info("No student background data is available.")
        return
    fields = [
        column
        for column in background.columns
        if background[column].nunique(dropna=True) <= 40
        and column not in {"id", "NO MATRIK", "created_at", "updated_at"}
    ]
    if not fields:
        st.dataframe(background, hide_index=True, width="stretch")
        return
    field = st.selectbox("Analyse field", fields)
    summary = (
        background[field]
        .fillna("Not specified")
        .astype(str)
        .value_counts()
        .rename_axis(field)
        .reset_index(name="Students")
    )
    a, b = st.columns(2)
    with a: metric("Total students", f"{len(background):,}", "Background records")
    with b: metric("Groups", f"{len(summary):,}", f"Distinct {field} values")
    st.dataframe(summary, hide_index=True, width="stretch")


def analysis_progress_page() -> None:
    heading(
        "Academic intelligence",
        "Analysis progress",
        "Review assessment marks and performance zones.",
    )
    progress, _ = fetch_table("stud_progress", DEMO_PROGRESS)
    assessments = [
        column
        for column in [
            "C1C2", "C5", "C8", "C9C10", "INDIVIDUAL ASSIGNMENT",
            "GROUP ASSIGNMENT", "UPS 1", "UPS 2", "UPS 3",
        ]
        if column in progress.columns
    ]
    if not assessments:
        assessments = [
            column for column in ["assignment", "pb", "task"]
            if column in progress.columns
        ]
    if progress.empty or not assessments:
        st.info("No assessment results are available.")
        return
    assessment = st.selectbox("Assessment", assessments)
    marks = pd.to_numeric(progress[assessment], errors="coerce")
    available = marks.dropna()
    a, b, c = st.columns(3)
    with a: metric("Results available", str(len(available)), f"of {len(progress)} students")
    with b: metric("Average mark", f"{available.mean():.1f}" if not available.empty else "—", assessment)
    with c: metric("Highest mark", f"{available.max():.1f}" if not available.empty else "—", assessment)
    zone_column = f"{assessment}_ZONE"
    columns = [c for c in ["NO MATRIK", assessment, zone_column] if c in progress.columns]
    st.dataframe(
        progress[columns].sort_values(assessment, ascending=False, na_position="last"),
        hide_index=True,
        width="stretch",
    )
    if zone_column in progress.columns:
        st.subheader("Zone distribution")
        zones = (
            progress[zone_column].fillna("Not assigned").astype(str)
            .value_counts().rename_axis("Zone").reset_index(name="Students")
        )
        st.dataframe(zones, hide_index=True, width="stretch")
    individuals, classes, _, _ = leaderboard_data()
    st.subheader("Individual progress leaderboard")
    individual_progress = individuals.dropna(subset=["Progress Average"]).sort_values(
        ["Progress Rank", "Student"]
    )[
        ["Progress Rank", "Student", "NO MATRIK", "Class", "Progress Average"]
    ]
    st.dataframe(individual_progress, hide_index=True, width="stretch")
    st.subheader("Class progress leaderboard")
    class_progress = classes.dropna(subset=["Progress Average"]).sort_values(
        ["Progress Rank", "Class"]
    )[
        ["Progress Rank", "Class", "Students", "Progress Average"]
    ]
    st.dataframe(class_progress, hide_index=True, width="stretch")


def analysis_xp_page() -> None:
    heading(
        "Recognition intelligence",
        "Analysis XP",
        "Review XP balances, badges and student rank.",
    )
    _, _, months, default_month = leaderboard_data()
    selected_month = st.selectbox(
        "Reward month",
        months,
        index=months.index(default_month),
    )
    individuals, classes, _, _ = leaderboard_data(selected_month)
    if individuals.empty:
        st.info("No XP records are available.")
        return
    individuals["Badge"] = pd.cut(
        individuals["Overall XP"],
        bins=[-1, 99, 299, 599, 999, float("inf")],
        labels=["None", "Starter", "Active learner", "Achiever", "XP Elite"],
    ).astype(str)
    a, b, c = st.columns(3)
    with a: metric("Monthly XP", f"{int(individuals['Monthly XP'].sum()):,}", selected_month)
    with b: metric("Overall XP", f"{int(individuals['Overall XP'].sum()):,}", "Cumulative balance")
    with c: metric("Ranked students", str(len(individuals)), "Eligible individuals")

    monthly_tab, overall_tab, class_tab = st.tabs(
        ["Monthly individual", "Overall individual", "Class leaderboard"]
    )
    with monthly_tab:
        monthly = individuals.sort_values(
            ["Monthly Rank", "Overall Rank", "Student"]
        )[
            [
                "Monthly Rank", "Student", "NO MATRIK", "Class",
                "Monthly XP", "Overall XP", "Badge",
            ]
        ]
        st.dataframe(monthly, hide_index=True, width="stretch")
    with overall_tab:
        overall = individuals.sort_values(["Overall Rank", "Student"])[
            [
                "Overall Rank", "Student", "NO MATRIK", "Class",
                "Overall XP", "Monthly XP", "Badge",
            ]
        ]
        st.dataframe(overall, hide_index=True, width="stretch")
    with class_tab:
        st.subheader(f"Monthly class ranking · {selected_month}")
        monthly_classes = classes.sort_values(["Monthly Rank", "Class"])[
            [
                "Monthly Rank", "Class", "Students", "Monthly XP Average",
                "Monthly XP Total", "Overall XP Average", "Class Badge",
            ]
        ]
        st.dataframe(monthly_classes, hide_index=True, width="stretch")
        st.subheader("Overall class ranking")
        overall_classes = classes.sort_values(["Overall Rank", "Class"])[
            [
                "Overall Rank", "Class", "Students", "Overall XP Average",
                "Overall XP Total", "Monthly XP Average", "Class Badge",
            ]
        ]
        st.dataframe(overall_classes, hide_index=True, width="stretch")


def analysis_material_page() -> None:
    heading(
        "Library intelligence",
        "Analysis material",
        "Analyse uploaded materials by chapter and resource type.",
    )
    materials, _ = fetch_table("materials", EMPTY_MATERIALS)
    if materials.empty:
        st.info("No materials have been uploaded yet.")
        return
    a, b, c = st.columns(3)
    with a: metric("Materials", str(len(materials)), "Uploaded resources")
    with b: metric("Chapters", str(materials.get("chapter", pd.Series()).nunique()), "With resources")
    with c: metric("Types", str(materials.get("material_type", pd.Series()).nunique()), "Resource classifications")
    if {"chapter", "material_type"}.issubset(materials.columns):
        summary = (
            materials.groupby(["chapter", "material_type"], dropna=False)
            .size().reset_index(name="Materials")
            .sort_values(["chapter", "material_type"])
        )
        st.dataframe(summary, hide_index=True, width="stretch")
    st.subheader("Material register")
    visible = [
        column for column in
        ["title", "course", "chapter", "material_type", "type", "uploaded_at"]
        if column in materials.columns
    ]
    st.dataframe(materials[visible], hide_index=True, width="stretch")


def admin_crud_page() -> None:
    heading(
        "Data management",
        "CRUD",
        "Add, update, delete or overwrite student datasets.",
    )
    target = st.selectbox(
        "Dataset",
        ["stud_background", "stud_progress", "stud_xp"],
    )
    data, live = fetch_table(target, pd.DataFrame())
    add_tab, update_tab, delete_tab, bulk_tab = st.tabs(
        ["Add", "Update", "Delete", "Bulk overwrite"]
    )
    with add_tab:
        st.caption("Add one record using the dataset columns below.")
        template_columns = [
            column for column in data.columns
            if column not in {"id", "created_at", "updated_at"}
        ]
        if not template_columns:
            st.info("Run the related Supabase schema before adding records.")
        else:
            with st.form(f"crud_add_{target}"):
                values = {
                    column: st.text_input(column)
                    for column in template_columns
                }
                if st.form_submit_button("Add record", type="primary"):
                    record = {key: value for key, value in values.items() if value != ""}
                    try:
                        db().table(target).insert(record).execute()
                        st.success("Record added.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Record could not be added: {exc}")
    with update_tab:
        if data.empty:
            st.info("No records are available to update.")
        else:
            edited = st.data_editor(data, hide_index=True, width="stretch")
            if st.button("Save updates", type="primary"):
                ok, message = upsert_rows(target, edited)
                (st.success if ok else st.error)(message)
    with delete_tab:
        if data.empty:
            st.info("No records are available to delete.")
        else:
            key = "NO MATRIK" if "NO MATRIK" in data.columns else "id"
            choices = data[key].dropna().astype(str).tolist()
            selected = st.selectbox("Record to delete", choices)
            confirmed = st.checkbox("I understand this permanently deletes the selected record.")
            if st.button("Delete record", disabled=not confirmed):
                try:
                    db().table(target).delete().eq(key, selected).execute()
                    st.success("Record deleted.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Record could not be deleted: {exc}")
    with bulk_tab:
        st.warning("Bulk overwrite replaces every record in the selected dataset.")
        uploaded = st.file_uploader(
            "CSV or Excel file", type=["csv", "xlsx", "xls"], key="crud_bulk_file"
        )
        if uploaded:
            frame = (
                pd.read_csv(uploaded)
                if uploaded.name.lower().endswith(".csv")
                else pd.read_excel(uploaded)
            )
            st.dataframe(frame.head(50), hide_index=True, width="stretch")
            confirmation = st.text_input("Type OVERWRITE to continue")
            if st.button("Overwrite dataset", disabled=confirmation != "OVERWRITE"):
                try:
                    key = "NO MATRIK"
                    db().table(target).delete().neq(key, "__never__").execute()
                    ok, message = upsert_rows(target, frame)
                    (st.success if ok else st.error)(message)
                except Exception as exc:
                    st.error(f"Overwrite stopped: {exc}")
    if not live:
        st.caption("The live dataset is unavailable or empty.")


def analytics_page() -> None:
    heading("Learning intelligence", "Performance analytics", "Turn class activity into clear, practical teaching decisions.")
    merged, _ = merged_students()
    numeric = [c for c in ["assignment", "pb", "task", "xp"] if c in merged]
    left, right = st.columns([1.4, 1])
    with left:
        if numeric:
            long = merged.melt(value_vars=numeric[:3], var_name="Assessment", value_name="Score")
            fig = px.box(long, x="Assessment", y="Score", color="Assessment", color_discrete_sequence=["#1F5DAA", "#17365D", "#8A96A3"])
            fig.update_layout(height=360, showlegend=False, margin=dict(l=5,r=5,t=20,b=5), plot_bgcolor="white", paper_bgcolor="white")
            polish_chart(fig)
            st.plotly_chart(fig, use_container_width=True)
    with right:
        st.subheader("Teaching signals")
        st.info("Published result performance is strongest in the leading assessment component.")
        st.warning("4 learners may need support in project-based assessment.")
        st.success("Task completion improved by 7% over the last four weeks.")
    if "xp" in merged:
        board = merged.sort_values("xp", ascending=False).reset_index(drop=True)
        board.insert(0, "rank", board.index + 1)
        st.subheader("Leaderboard")
        keep = [c for c in ["rank", "name", "student_name", "student_id", "xp", "badge"] if c in board]
        st.dataframe(board[keep], use_container_width=True, hide_index=True)


def admin_dashboard() -> None:
    heading("Administration", "XPLMS control centre", "Manage people, data quality and platform health with confidence.")
    merged, live = merged_students()
    a, b, c, d = st.columns(4)
    with a: metric("Student records", f"{len(merged):,}", "Active learners")
    with b: metric("Admin users", "16", "Across 6 programmes")
    with c: metric("Data health", "98.4%", "Last sync 12 min ago")
    with d: metric("Platform status", "Healthy", "All services normal")
    st.subheader("Quick actions")
    cols = st.columns(3)
    cols[0].button("＋ Add student", use_container_width=True)
    cols[1].button("⇧ Import dataset", use_container_width=True)
    cols[2].button("◉ Manage access", use_container_width=True)
    if not live: st.warning("Supabase tables are not currently available to the app. Apply the included database schema and verify your project URL/key.")


def user_access_page() -> None:
    heading(
        "Administration",
        "User access",
        "Create and maintain XPLMS admin and student accounts.",
    )
    client = db()
    if not client:
        st.error("The server connection is unavailable.")
        return
    users = pd.DataFrame(
        client.table("app_users")
        .select('id,email,full_name,role,"NO MATRIK",active,must_change_password,last_login_at,created_at')
        .order("full_name")
        .execute()
        .data
    )
    create_tab, manage_tab, bulk_tab = st.tabs(
        ["Add user", "Manage users", "Bulk import"]
    )
    with create_tab:
        role = st.segmented_control("Access level", ["Admin", "Student"], default="Student")
        students, _ = fetch_table("stud_background", DEMO_STUDENTS)
        matric = None
        suggested_name = ""
        can_create = True
        if role == "Student":
            matric_col = "NO MATRIK" if "NO MATRIK" in students.columns else "student_id"
            name_col = "NAMA PELAJAR" if "NAMA PELAJAR" in students.columns else "name"
            assigned = set(users.get("NO MATRIK", pd.Series(dtype=str)).dropna().astype(str))
            available = students[~students[matric_col].astype(str).isin(assigned)].copy()
            options = available[matric_col].astype(str).tolist()
            if not options:
                st.info("Every student record is already linked to an account.")
                can_create = False
            else:
                name_map = dict(zip(available[matric_col].astype(str), available[name_col].astype(str)))
                matric = st.selectbox("Student record", options, format_func=lambda value: f"{name_map[value]} · {value}")
                suggested_name = name_map[matric]
        with st.form("create_app_user", clear_on_submit=True):
            full_name = st.text_input("Full name", value=suggested_name)
            email = st.text_input("Email address")
            temp_password = st.text_input("Temporary password", type="password")
            if st.form_submit_button(
                "Create user", type="primary", disabled=not can_create
            ):
                if not full_name.strip() or "@" not in email:
                    st.error("Enter a valid name and email.")
                elif not valid_password(temp_password):
                    st.error("Use at least 10 characters with uppercase, lowercase and a number.")
                else:
                    try:
                        client.table("app_users").insert({
                            "email": email.strip().lower(),
                            "full_name": full_name.strip(),
                            "role": role.lower(),
                            "NO MATRIK": matric if role == "Student" else None,
                            "password_hash": hash_password(temp_password),
                            "active": True,
                            "must_change_password": True,
                        }).execute()
                        st.success(f"{role} account created. Share the temporary password securely.")
                    except Exception as exc:
                        st.error(f"User could not be created: {exc}")
    with manage_tab:
        if users.empty:
            st.info("No users found.")
            return
        st.dataframe(
            users.drop(columns=["id"], errors="ignore"),
            hide_index=True,
            width="stretch",
        )
        user_ids = users["id"].astype(str).tolist()
        label_map = {
            str(row["id"]): f"{row['full_name']} · {row['email']} · {row['role']}"
            for _, row in users.iterrows()
        }
        selected = st.selectbox("Select account", user_ids, format_func=lambda value: label_map[value])
        selected_row = users[users["id"].astype(str) == selected].iloc[0]
        c1, c2 = st.columns(2)
        active = c1.toggle("Account active", value=bool(selected_row["active"]))
        if c1.button("Save account status"):
            if selected == str(st.session_state.user["id"]) and not active:
                st.error("You cannot deactivate your own active session.")
            else:
                client.table("app_users").update({"active": active}).eq("id", selected).execute()
                st.success("Account status updated.")
        new_password = c2.text_input("New temporary password", type="password")
        if c2.button("Reset password"):
            if not valid_password(new_password):
                st.error("Use at least 10 characters with uppercase, lowercase and a number.")
            else:
                client.table("app_users").update({
                    "password_hash": hash_password(new_password),
                    "must_change_password": True,
                }).eq("id", selected).execute()
                st.success("Temporary password set. The user must change it at next login.")
        st.divider()
        remove_confirmed = st.checkbox(
            "Permanently remove this account",
            key="remove_user_confirmed",
        )
        if st.button("Remove user", disabled=not remove_confirmed):
            if selected == str(st.session_state.user["id"]):
                st.error("You cannot remove your own active account.")
            else:
                client.table("app_users").delete().eq("id", selected).execute()
                st.success("User account removed.")
                st.rerun()
    with bulk_tab:
        st.write(
            "Upload CSV or Excel with: email, full_name, role, NO MATRIK and "
            "temporary_password. Role must be admin or student."
        )
        upload = st.file_uploader(
            "User access file",
            type=["csv", "xlsx", "xls"],
            key="user_access_bulk_file",
        )
        if upload:
            frame = (
                pd.read_csv(upload)
                if upload.name.lower().endswith(".csv")
                else pd.read_excel(upload)
            )
            st.dataframe(frame.head(50), hide_index=True, width="stretch")
            required = {"email", "full_name", "role", "temporary_password"}
            missing = required - set(frame.columns)
            if missing:
                st.error(f"Missing columns: {', '.join(sorted(missing))}")
            elif st.button("Import user access", type="primary"):
                try:
                    records = []
                    for _, row in frame.iterrows():
                        role_value = str(row["role"]).strip().lower()
                        if role_value not in {"admin", "student"}:
                            raise ValueError(f"Invalid role: {row['role']}")
                        password = str(row["temporary_password"])
                        if not valid_password(password):
                            raise ValueError(
                                f"Temporary password for {row['email']} is not strong enough."
                            )
                        matric = row.get("NO MATRIK")
                        records.append({
                            "email": str(row["email"]).strip().lower(),
                            "full_name": str(row["full_name"]).strip(),
                            "role": role_value,
                            "NO MATRIK": (
                                str(matric).strip()
                                if role_value == "student" and pd.notna(matric)
                                else None
                            ),
                            "password_hash": hash_password(password),
                            "active": True,
                            "must_change_password": True,
                        })
                    client.table("app_users").upsert(
                        records, on_conflict="email"
                    ).execute()
                    st.success(f"{len(records)} user accounts imported.")
                except Exception as exc:
                    st.error(f"User import failed: {exc}")


def change_password_page() -> None:
    st.title("Set a new password")
    st.write("Your administrator issued a temporary password. Replace it before continuing.")
    with st.form("change_password"):
        password = st.text_input("New password", type="password")
        confirm = st.text_input("Confirm new password", type="password")
        if st.form_submit_button("Update password", type="primary"):
            if password != confirm:
                st.error("The passwords do not match.")
            elif not valid_password(password):
                st.error("Use at least 10 characters with uppercase, lowercase and a number.")
            else:
                client = db()
                client.table("app_users").update({
                    "password_hash": hash_password(password),
                    "must_change_password": False,
                }).eq("id", st.session_state.user["id"]).execute()
                st.session_state.user["must_change_password"] = False
                st.success("Password updated.")
                st.rerun()


def import_page() -> None:
    heading("Data operations", "Bulk import", "Review changes before replacing or updating a dataset.")
    target = st.selectbox("Target dataset", ["stud_background", "stud_progress"])
    uploaded = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx", "xls"])
    if uploaded:
        try:
            frame = pd.read_csv(uploaded) if uploaded.name.lower().endswith(".csv") else pd.read_excel(uploaded)
            st.success(f"{len(frame)} records ready · {len(frame.columns)} columns")
            st.dataframe(frame.head(50), use_container_width=True, hide_index=True)
            mode = st.radio("Import method", ["Update matching records", "Replace entire dataset"], horizontal=True)
            confirm = st.checkbox("I have reviewed this preview and want to continue")
            if st.button("Import records", type="primary", disabled=not confirm):
                client = db()
                if mode.startswith("Replace") and client:
                    try:
                        client.table(target).delete().neq("id", "__never__").execute()
                    except Exception as exc:
                        st.error(f"Replace was stopped safely: {exc}")
                        return
                ok, message = upsert_rows(target, frame)
                (st.success if ok else st.error)(message)
        except Exception as exc:
            st.error(f"Could not read this file: {exc}")
    st.download_button("Download student CSV template", DEMO_STUDENTS.head(0).to_csv(index=False), "student_import_template.csv", "text/csv")


def simple_page(title: str, copy: str) -> None:
    heading("XPLMS", title, copy)
    st.info("This workspace is ready for your institution-specific configuration.")


def main() -> None:
    demo_role = os.getenv("XPLMS_DEMO_ROLE", "").title()
    if "user" not in st.session_state and demo_role in {"Student", "Admin"}:
        st.session_state.user = {
            "name": f"Demo {demo_role}",
            "email": "demo@xplms.edu",
            "role": demo_role,
            "id": "demo",
        }
    if "user" not in st.session_state:
        login()
        return
    user = st.session_state.user
    if user.get("must_change_password"):
        change_password_page()
        return
    actual_role = user.get("role", "Student")
    role = (
        st.session_state.get("admin_preview_role", "Admin")
        if actual_role == "Admin" and not is_demo()
        else actual_role
    )
    if role not in {"Student", "Admin"}:
        role = "Student"
    display_name = student_nickname(user) if role == "Student" else user["name"]
    page = sidebar(role, display_name)
    if role == "Student":
        {
            "Progress": progress_page,
            "Materials": materials_page,
            "Quiz": quiz_page,
            "Leaderboard": student_leaderboard_page,
            "Profile": profile_page,
            "XP": xp_badge_page,
        }[page]()
    else:
        {
            "Student record": admin_student_records_page,
            "Analysis background": analysis_background_page,
            "Analysis progress": analysis_progress_page,
            "Analysis XP": analysis_xp_page,
            "Analysis material": analysis_material_page,
            "CRUD": admin_crud_page,
            "Award XP": award_xp_page,
            "User access": user_access_page,
            "Material": lambda: materials_page(True),
            "Quiz": admin_quiz_page,
        }[page]()


if __name__ == "__main__":
    main()
