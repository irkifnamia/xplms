"""XPLMS — Experience Point Learning Management System."""

from __future__ import annotations

import io
import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any

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
    questions: list[GeneratedQuestion] = Field(min_length=3, max_length=15)


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
    [data-testid="stHeader"],
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

    [data-testid="stAppViewContainer"] .block-container {
        padding-top: 1rem !important;
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
  margin-left:auto !important;
  margin-right:auto !important;
  text-align:center !important;
}
[data-testid="stImage"] img {
  display:block !important;
  margin-left:auto !important;
  margin-right:auto !important;
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
    height:0 !important;
    min-height:0 !important;
    background:transparent !important;
  }
  [data-testid="stHeader"] [data-testid="stToolbar"],
  [data-testid="stHeader"] [data-testid="stToolbarActions"],
  [data-testid="stHeader"] [data-testid="stStatusWidget"],
  [data-testid="stHeader"] [data-testid="stDeployButton"] {
    display:none !important;
  }
  [data-testid="collapsedControl"] {
    display:flex !important;
    visibility:visible !important;
    position:fixed !important;
    top:.7rem !important;
    left:.7rem !important;
    z-index:1000000 !important;
  }
  [data-testid="collapsedControl"] button {
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
  [data-testid="collapsedControl"] svg {
    color:#17233a !important;
    fill:#17233a !important;
  }
  body:has(.student-sidebar-marker) [data-testid="collapsedControl"] {
    display:none !important;
  }
  body:has(.student-sidebar-marker) .block-container {
    padding:0.65rem 0.72rem 8.25rem !important;
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

  /* Student navigation becomes a thumb-friendly fixed bottom bar. */
  section[data-testid="stSidebar"]:has(.student-sidebar-marker) {
    display:block !important;
    position:fixed !important;
    inset:auto 0 0 0 !important;
    width:100vw !important;
    min-width:100vw !important;
    height:auto !important;
    z-index:999999 !important;
    transform:none !important;
    border:0 !important;
    border-top:1px solid #cbd6e2 !important;
    box-shadow:0 -6px 20px rgba(23,35,58,.12);
  }
  section[data-testid="stSidebar"]:has(.student-sidebar-marker)
  [data-testid="stSidebarContent"] {
    padding:.38rem .45rem .45rem !important;
    background:#fff !important;
  }
  section[data-testid="stSidebar"]:has(.student-sidebar-marker)
  [data-testid="stSidebarContent"] > div > div > div > :is(
    [data-testid="stImage"],
    [data-testid="stCaptionContainer"],
    [data-testid="stMarkdownContainer"],
    [data-testid="stSelectbox"],
    [data-testid="stDivider"]
  ) { display:none !important; }
  section[data-testid="stSidebar"]:has(.student-sidebar-marker)
  [data-testid="stRadio"] [role="radiogroup"] {
    display:flex !important;
    flex-direction:row !important;
    gap:2px !important;
    width:100% !important;
  }
  section[data-testid="stSidebar"]:has(.student-sidebar-marker)
  [data-testid="stRadio"] label {
    flex:1 1 20%;
    min-width:0;
    min-height:48px;
    margin:0 !important;
    padding:.36rem .12rem !important;
    display:flex !important;
    justify-content:center;
    text-align:center;
    line-height:1.12;
    font-size:.69rem;
  }
  section[data-testid="stSidebar"]:has(.student-sidebar-marker)
  [data-testid="stRadio"] label > div:first-child {
    display:none !important;
  }
  section[data-testid="stSidebar"]:has(.student-sidebar-marker)
  [data-testid="stRadio"] label:has(input:checked) {
    background:#e8f1fc !important;
    color:#154b8c !important;
  }
  section[data-testid="stSidebar"]:has(.student-sidebar-marker)
  .stButton > button {
    min-height:2rem !important;
    margin-top:.2rem;
    font-size:.72rem;
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
        response = OpenAI(api_key=api_key).responses.parse(
            model="gpt-5.6-luna",
            input=[
                {
                    "role": "system",
                    "content": (
                        "Create a fair educational multiple-choice quiz using only the supplied "
                        "study material. Use four plausible options per question, one correct answer, "
                        "clear explanations, and no trick questions."
                    ),
                },
                {
                    "role": "user",
                    "content": f"Create exactly {question_count} questions from this material:\n\n{source_text}",
                },
            ],
            text_format=GeneratedQuiz,
        )
        generated = response.output_parsed
        quiz_row = client.table("quizzes").insert({
            "material_id": int(material["id"]),
            "title": generated.title,
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
            for index, item in enumerate(generated.questions, start=1)
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
            "Student": ["Progress", "Materials", "Quiz", "Profile", "XP & Badge"],
            "Admin": [
                "Dashboard", "Award XP", "Student records", "Materials", "Analytics",
                "User access", "Data import", "System",
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
    return merged, live_a and live_b and live_c


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

    name = row.get("NAMA PELAJAR", row.get("name", user.get("name", "—")))
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
    heading("Recognition", "XP & Badge", "Track your experience points, badge progress and class rank.")
    user = st.session_state.get("user", {})
    xp_data, _ = fetch_table("stud_xp", DEMO_PROGRESS[["student_id", "xp"]].copy())
    matric_col = "NO MATRIK" if "NO MATRIK" in xp_data.columns else "student_id"
    xp_col = "XP" if "XP" in xp_data.columns else "xp"
    if user.get("no_matrik") and matric_col == "NO MATRIK":
        own = xp_data[xp_data[matric_col] == user["no_matrik"]]
    else:
        own = xp_data.head(1)
    if own.empty:
        st.info("No XP record is linked to this account.")
        return
    xp_value = pd.to_numeric(pd.Series([own.iloc[0].get(xp_col, 0)]), errors="coerce").fillna(0).iloc[0]
    ranked = xp_data.copy()
    ranked["_xp"] = pd.to_numeric(ranked[xp_col], errors="coerce").fillna(0)
    ranked = ranked.sort_values(["_xp", matric_col], ascending=[False, True]).reset_index(drop=True)
    ranked["Rank"] = ranked.index + 1
    own_matric = str(own.iloc[0][matric_col])
    rank_rows = ranked[ranked[matric_col].astype(str) == own_matric]
    rank = int(rank_rows.iloc[0]["Rank"]) if not rank_rows.empty else 0

    badge_rules = [
        (100, "Starter", "Begin earning experience points"),
        (300, "Active learner", "Build a consistent learning record"),
        (600, "Achiever", "Demonstrate sustained participation"),
        (1000, "XP Elite", "Reach an outstanding XP milestone"),
    ]
    earned = [badge for threshold, badge, _ in badge_rules if xp_value >= threshold]
    next_badge = next(((threshold, badge) for threshold, badge, _ in badge_rules if xp_value < threshold), None)
    a, b, c = st.columns(3)
    with a: metric("Total XP", f"{int(xp_value):,}", "Current experience balance")
    with b: metric("Class rank", f"#{rank}" if rank else "—", f"of {len(ranked)} students")
    with c: metric("Badges earned", str(len(earned)), f"of {len(badge_rules)} available")

    st.subheader("Badge collection")
    cols = st.columns(len(badge_rules))
    for col, (threshold, badge, description) in zip(cols, badge_rules):
        status = "Unlocked" if xp_value >= threshold else f"{max(0, threshold-int(xp_value))} XP to unlock"
        with col:
            st.markdown(
                f'<div class="metric-card" style="text-align:center"><div class="metric-value">{threshold}</div>'
                f'<b>{badge}</b><div class="metric-note">{description}<br>{status}</div></div>',
                unsafe_allow_html=True,
            )
    if next_badge:
        threshold, badge = next_badge
        st.progress(min(float(xp_value) / threshold, 1.0), text=f"Progress to {badge}: {int(xp_value)} / {threshold} XP")
    else:
        st.success("All current XP badges unlocked.")

    with st.expander("Submit mathematics discussion proof"):
        st.write("Share evidence of a meaningful mathematics discussion for Admin review.")
        with st.form("xp_proof_claim", clear_on_submit=True):
            title = st.text_input("Discussion topic")
            description = st.text_area("What did you discuss or learn?")
            proof = st.file_uploader("Proof image", type=["jpg", "jpeg", "png", "webp"])
            if st.form_submit_button("Send for approval", type="primary"):
                if not title.strip() or not description.strip() or not proof:
                    st.error("Complete the topic, description and proof image.")
                elif is_demo():
                    st.success("Demo claim submitted for Admin approval.")
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
                                "claim_type": "math_discussion",
                                "title": title.strip(),
                                "description": description.strip(),
                                "proof_path": path,
                                "status": "pending",
                            }).execute()
                            st.success("Proof submitted. An Admin will review your claim.")
                        except Exception as exc:
                            st.error(f"Claim could not be submitted: {exc}")

    events_fallback = pd.DataFrame(
        [{"rule_code": "consultation", "points": 20, "reason": "Academic consultation", "created_at": "2026-07-28"}]
    )
    events, live = fetch_table("xp_events", events_fallback)
    if live and "NO MATRIK" in events.columns:
        events = events[events["NO MATRIK"] == user.get("no_matrik")]
    st.subheader("XP history")
    if events.empty:
        st.info("No XP awards recorded yet.")
    else:
        visible = [c for c in ["created_at", "rule_code", "reason", "points"] if c in events.columns]
        st.dataframe(events[visible].sort_values("created_at", ascending=False), hide_index=True, width="stretch")


def materials_page(lecturer: bool = False) -> None:
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

    question_count = 5
    if lecturer:
        question_count = st.number_input("Questions per generated quiz", 3, 15, 5)

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
            if lecturer and live and item.get("file_path") and c4.button("Generate quiz", key=f"quiz_material_{i}"):
                with st.spinner("Generating a draft quiz from this material…"):
                    ok, message = generate_material_quiz(item, int(question_count))
                (st.success if ok else st.error)(message)

    if lecturer:
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


def quiz_page() -> None:
    heading("Knowledge check", "In-app quizzes", "Complete material-based quizzes and earn XP automatically.")
    quizzes_fallback = pd.DataFrame(
        [{"id": 1, "title": "Algebra foundations", "instructions": "Choose the best answer.", "xp_reward": 25, "passing_score": 60, "status": "published"}]
    )
    quizzes, live = fetch_table("quizzes", quizzes_fallback)
    if "status" in quizzes.columns:
        quizzes = quizzes[quizzes["status"] == "published"]
    if quizzes.empty:
        st.info("No quizzes are currently available.")
        return
    user = st.session_state.user
    for _, quiz in quizzes.iterrows():
        with st.expander(f"{quiz['title']} · {quiz.get('xp_reward', 0)} XP"):
            st.write(quiz.get("instructions", "Choose the best answer."))
            if live:
                previous = db().table("quiz_attempts").select("score,passed,completed_at").eq("quiz_id", quiz["id"]).eq("student_user_id", user["id"]).limit(1).execute().data
                if previous:
                    result = previous[0]
                    st.success(f"Completed · Score {float(result['score']):.0f}% · {'Passed' if result['passed'] else 'Not passed'}")
                    continue
                questions = db().table("quiz_questions").select("*").eq("quiz_id", quiz["id"]).order("position").execute().data
            else:
                questions = [
                    {"id": 1, "question": "What is 2 + 3?", "options": ["4", "5", "6", "7"], "correct_index": 1, "explanation": "2 + 3 equals 5."},
                    {"id": 2, "question": "Which number is even?", "options": ["3", "5", "8", "9"], "correct_index": 2, "explanation": "8 is divisible by 2."},
                    {"id": 3, "question": "What is 4 × 2?", "options": ["6", "7", "8", "9"], "correct_index": 2, "explanation": "Four groups of two equal 8."},
                ]
            if not questions:
                st.warning("This quiz has no questions.")
                continue
            with st.form(f"quiz_attempt_{quiz['id']}"):
                answers = {}
                for question in questions:
                    options = question["options"]
                    if isinstance(options, str):
                        options = json.loads(options)
                    answers[str(question["id"])] = st.radio(
                        question["question"],
                        range(len(options)),
                        format_func=lambda index, opts=options: opts[index],
                        index=None,
                        key=f"answer_{quiz['id']}_{question['id']}",
                    )
                if st.form_submit_button("Submit quiz", type="primary"):
                    if any(answer is None for answer in answers.values()):
                        st.error("Answer every question before submitting.")
                    else:
                        correct = sum(
                            int(answers[str(question["id"])]) == int(question["correct_index"])
                            for question in questions
                        )
                        score = correct / len(questions) * 100
                        passed = score >= float(quiz.get("passing_score", 60))
                        if not live:
                            (st.success if passed else st.warning)(f"Demo score: {score:.0f}%")
                        else:
                            attempt = db().table("quiz_attempts").insert({
                                "quiz_id": quiz["id"],
                                "student_user_id": user["id"],
                                "NO MATRIK": user["no_matrik"],
                                "answers": answers,
                                "score": score,
                                "correct_count": correct,
                                "total_questions": len(questions),
                                "passed": passed,
                            }).execute().data[0]
                            xp_awarded = 0
                            if passed and int(quiz.get("xp_reward", 0)) > 0:
                                xp_awarded = int(quiz["xp_reward"])
                                event = db().table("xp_events").insert({
                                    "NO MATRIK": user["no_matrik"],
                                    "rule_code": "quiz_completion",
                                    "points": xp_awarded,
                                    "source_id": f"quiz-{quiz['id']}",
                                    "reason": f"Passed quiz: {quiz['title']}",
                                    "award_mode": "automatic",
                                    "awarded_by": None,
                                }).execute().data[0]
                                db().table("quiz_attempts").update({"xp_event_id": event["id"]}).eq("id", attempt["id"]).execute()
                            if passed:
                                st.success(f"Passed with {score:.0f}%. {xp_awarded} XP awarded.")
                            else:
                                st.warning(f"Score: {score:.0f}%. The passing score is {quiz.get('passing_score', 60)}%.")
                            st.rerun()


def award_xp_page() -> None:
    heading(
        "Recognition",
        "Award experience points",
        "Recognise academic consultation and meaningful class participation.",
    )
    students, students_live = merged_students()
    rules_fallback = pd.DataFrame(
        [
            {"code": "consultation", "name": "Consultation", "default_points": 20, "award_mode": "manual"},
            {"code": "class_participation", "name": "Class participation", "default_points": 10, "award_mode": "manual"},
        ]
    )
    rules, rules_live = fetch_table("xp_rules", rules_fallback)
    if "award_mode" in rules.columns:
        rules = rules[rules["award_mode"] == "manual"]
    if "code" in rules.columns:
        rules = rules[rules["code"].isin(["consultation", "class_participation"])]
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

    st.subheader("Mathematics discussion claims")
    claims_fallback = pd.DataFrame(
        [{
            "id": 1, "NO MATRIK": "S24001", "title": "Algebra discussion",
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
            with st.expander(f"{claim.get('title', 'Math discussion')} · {claim.get('NO MATRIK', '')}"):
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
                            rule = client.table("xp_rules").select("default_points").eq("code", "math_discussion").limit(1).execute().data
                            points = int(rule[0]["default_points"]) if rule else 20
                            event = client.table("xp_events").insert({
                                "NO MATRIK": claim["NO MATRIK"],
                                "rule_code": "math_discussion",
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
        "- **Admin awards:** consultation and class participation.\n"
        "- **Admin-approved claims:** student mathematics discussion with image proof.\n"
        "- **Automatic awards:** successful in-app material quizzes."
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
    create_tab, manage_tab = st.tabs(["Add user", "Manage users"])
    with create_tab:
        role = st.segmented_control("Access level", ["Admin", "Student"], default="Student")
        students, _ = fetch_table("stud_background", DEMO_STUDENTS)
        matric = None
        suggested_name = ""
        if role == "Student":
            matric_col = "NO MATRIK" if "NO MATRIK" in students.columns else "student_id"
            name_col = "NAMA PELAJAR" if "NAMA PELAJAR" in students.columns else "name"
            assigned = set(users.get("NO MATRIK", pd.Series(dtype=str)).dropna().astype(str))
            available = students[~students[matric_col].astype(str).isin(assigned)].copy()
            options = available[matric_col].astype(str).tolist()
            if not options:
                st.info("Every student record is already linked to an account.")
                return
            name_map = dict(zip(available[matric_col].astype(str), available[name_col].astype(str)))
            matric = st.selectbox("Student record", options, format_func=lambda value: f"{name_map[value]} · {value}")
            suggested_name = name_map[matric]
        with st.form("create_app_user", clear_on_submit=True):
            full_name = st.text_input("Full name", value=suggested_name)
            email = st.text_input("Email address")
            temp_password = st.text_input("Temporary password", type="password")
            if st.form_submit_button("Create user", type="primary"):
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
    page = sidebar(role, user["name"])
    if role == "Student":
        {
            "Progress": progress_page,
            "Materials": materials_page,
            "Quiz": quiz_page,
            "Profile": profile_page,
            "XP & Badge": xp_badge_page,
        }[page]()
    else:
        {
            "Dashboard": admin_dashboard,
            "Award XP": award_xp_page,
            "Student records": lambda: students_page(True),
            "Materials": lambda: materials_page(True),
            "Analytics": analytics_page,
            "User access": user_access_page,
            "Data import": import_page,
            "System": lambda: simple_page(
                "System settings", "Configure programmes, XP rules, grading and integrations."
            ),
        }[page]()


if __name__ == "__main__":
    main()
