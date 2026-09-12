
import os
import io
import re
import html
import textwrap
from datetime import datetime

import streamlit as st
import pandas as pd
from groq import Groq
from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors
from docx import Document
from docx.shared import Pt, Inches

from langchain_community.document_loaders import TextLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_community.vectorstores import FAISS


# ============================================================
# APP CONFIG
# ============================================================

st.set_page_config(
    page_title="CareNote AI | Discharge Summary",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

MODEL_NAME = "openai/gpt-oss-20b"
DISCLAIMER = (
    "AI-generated draft only. This tool supports documentation and does not "
    "diagnose, prescribe, modify medication/dosage, or replace clinical judgment. "
    "A qualified healthcare professional must review and approve every summary."
)


# ============================================================
# CARENOTE AI — MODERN UI
# ============================================================

st.markdown(
    """
    <style>
    :root {
        --navy: #10264d;
        --navy-2: #172f5f;
        --blue: #4f7cff;
        --cyan: #38c7d6;
        --purple: #7858e8;
        --text: #19345d;
        --muted: #71819a;
        --surface: rgba(255,255,255,.88);
        --line: #dfe8f5;
    }

    .stApp {
        background:
            radial-gradient(circle at 85% 8%, rgba(104,151,255,.18), transparent 28%),
            radial-gradient(circle at 12% 85%, rgba(72,211,209,.15), transparent 30%),
            linear-gradient(135deg, #f5f9ff 0%, #eef7ff 48%, #f9fbff 100%);
    }

    .block-container {
        max-width: 1450px;
        padding-top: 1.25rem;
        padding-bottom: 3rem;
    }

    #MainMenu, footer { visibility: hidden; }
    header { background: transparent !important; }

    h1,h2,h3,h4 { color: var(--text); letter-spacing: -.025em; }
    p,li,label { color: #4f6079; }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #10264d 0%, #152f60 62%, #101f43 100%);
        border: none;
    }
    section[data-testid="stSidebar"] > div {
        background: transparent;
    }
    section[data-testid="stSidebar"] .block-container {
        padding: 1.5rem 1rem 1.25rem;
    }
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] .stCaption {
        color: #c8d6ee !important;
    }

    .brand-wrap { padding: 8px 8px 22px; }
    .brand-mark {
        width: 48px; height: 48px; border-radius: 16px;
        display: inline-flex; align-items:center; justify-content:center;
        background: linear-gradient(135deg,#42d9df,#7858e8);
        color: white; font-size: 25px; box-shadow: 0 10px 25px rgba(56,199,214,.22);
        vertical-align: middle; margin-right: 10px;
    }
    .brand-name { color: white; font-size: 22px; font-weight: 850; vertical-align: middle; }
    .brand-ai { color: #88a4ff; }
    .brand-sub { color:#a9bad9; font-size:11px; margin:7px 0 0 58px; }

    .sidebar-label { color:#8ea5ca; font-size:10px; font-weight:800; text-transform:uppercase; letter-spacing:.13em; margin:18px 8px 8px; }

    .topbar {
        display:flex; align-items:center; justify-content:space-between;
        padding: 8px 2px 20px;
    }
    .topbar-title { font-size:14px; font-weight:800; color:#24416e; }
    .user-chip {
        display:flex; align-items:center; gap:10px; padding:8px 13px;
        border-radius:999px; background:rgba(255,255,255,.78); border:1px solid #dce7f6;
        box-shadow:0 5px 18px rgba(36,65,110,.06); color:#24416e; font-size:13px; font-weight:700;
    }
    .avatar { width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:#eaf1ff;color:#496dd6; }

    .hero {
        position:relative; overflow:hidden; background:rgba(255,255,255,.88);
        border:1px solid rgba(213,226,244,.9); border-radius:30px; padding:42px 46px;
        box-shadow:0 18px 55px rgba(38,70,115,.09); margin-bottom:22px;
    }
    .hero:after {
        content:""; position:absolute; width:310px;height:310px;border-radius:50%;right:-110px;top:-135px;
        background:linear-gradient(135deg,rgba(78,126,255,.16),rgba(71,210,207,.06));
    }
    .eyebrow { color:#5a72c9; font-size:11px; font-weight:850; letter-spacing:.14em; text-transform:uppercase; }
    .hero h1 { font-size:43px; margin:7px 0 9px; color:#132f59; }
    .hero p { font-size:16px; max-width:760px; margin:0; color:#71819a; line-height:1.65; }
    .hero-badges { margin-top:19px; display:flex; gap:8px; flex-wrap:wrap; }
    .badge { display:inline-block; padding:7px 12px; border-radius:999px; background:#eef4ff; color:#496ac1; font-size:11px; font-weight:800; border:1px solid #dce7fb; }

    .welcome-card {
        background:rgba(255,255,255,.93); border:1px solid #dce7f4; border-radius:30px;
        padding:45px 40px; box-shadow:0 18px 60px rgba(38,70,115,.10); text-align:center;
    }
    .welcome-icon { width:76px;height:76px;border-radius:26px;margin:0 auto 18px;display:flex;align-items:center;justify-content:center;font-size:39px;color:white;background:linear-gradient(135deg,#3bcbd5,#7656e8);box-shadow:0 14px 30px rgba(90,88,225,.20); }
    .welcome-card h1 { font-size:40px; margin:0; }
    .welcome-card .lead { color:#6d7d96; font-size:15px; margin:8px 0 26px; }

    .feature-panel { background:rgba(255,255,255,.65); border:1px solid #dce7f4; border-radius:28px; padding:22px; }
    .feature-photo {
        height:190px; border-radius:22px; margin-bottom:18px;
        background:linear-gradient(135deg,#dceeff,#eef7ff 50%,#e7ecff);
        display:flex; align-items:center; justify-content:center; font-size:70px;
    }
    .feature-item { display:flex; gap:13px; padding:13px 4px; }
    .feature-dot { width:40px;height:40px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:#edf3ff;font-size:20px;flex:0 0 auto; }
    .feature-item b { color:#203f72; font-size:13px; }
    .feature-item span { display:block; color:#7a8aa2; font-size:11px; margin-top:3px; line-height:1.4; }

    .security-box { margin-top:20px; padding:13px 15px; border-radius:16px; background:#f1f6ff; border:1px solid #dce8fb; color:#57709b; font-size:11px; text-align:left; }

    .card {
        background:rgba(255,255,255,.88); border:1px solid #dce7f4; border-radius:22px;
        padding:22px; box-shadow:0 9px 30px rgba(38,70,115,.055); margin-bottom:16px;
    }
    .card-title { font-size:17px; font-weight:850; color:#1b3b6b; margin-bottom:4px; }
    .card-sub { font-size:12px; color:#8190a7; margin-bottom:15px; }

    .metric { background:rgba(255,255,255,.82); border:1px solid #dce7f4; border-radius:18px; padding:17px; box-shadow:0 7px 25px rgba(38,70,115,.045); }
    .metric .label { font-size:10px; font-weight:800; color:#7d8da7; text-transform:uppercase; letter-spacing:.1em; }
    .metric .value { color:#183861; font-size:24px; font-weight:850; margin-top:5px; }

    .notice { padding:15px 17px; border-radius:16px; background:#f0f7ff; border:1px solid #d8e8fa; color:#536b8d; font-size:12px; line-height:1.55; }
    .warning { padding:15px 17px; border-radius:16px; background:#fff8e7; border:1px solid #f1dfb3; color:#765e25; font-size:12px; line-height:1.55; }
    .success-box { padding:15px 17px; border-radius:16px; background:#effaf7; border:1px solid #d2eee7; color:#397468; font-size:12px; }

    .stButton > button { border-radius:13px; min-height:43px; font-weight:800; border:1px solid #d9e5f4; }
    .stButton > button[kind="primary"] { background:linear-gradient(135deg,#385bd6,#7858e8); border:none; color:white; }
    .stDownloadButton > button { border-radius:13px; font-weight:800; }
    [data-testid="stFileUploader"] { border-radius:18px; background:rgba(255,255,255,.8); }
    textarea, input { border-radius:12px !important; }
    div[data-testid="stExpander"] { border:1px solid #dce7f4; border-radius:16px; background:rgba(255,255,255,.7); }
    .footer-note { margin-top:35px; padding:17px 2px; border-top:1px solid #dbe5f2; color:#8090a7; font-size:10px; line-height:1.6; }

    /* Sidebar radio navigation */
    section[data-testid="stSidebar"] div[role="radiogroup"] label {
        background:transparent !important; border-radius:12px; padding:8px 10px !important; margin:3px 0 !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label:hover { background:rgba(255,255,255,.08) !important; }
    section[data-testid="stSidebar"] div[role="radiogroup"] p { color:#d9e5f8 !important; font-weight:700; }
    section[data-testid="stSidebar"] div[role="radiogroup"] label[data-checked="true"] { background:linear-gradient(90deg,rgba(92,112,231,.9),rgba(91,91,210,.55)) !important; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SESSION STATE
# ============================================================

DEFAULTS = {
    "patient_data": "",
    "discharge_summary": "",
    "edited_summary": "",
    "qa_answer": "",
    "qa_question": "",
    "rag_sources": [],
    "missing_fields": [],
    "grounding_result": "",
    "uploaded_names": [],
    "language": "English",
    "section_preview": "",
    "patient_message": "",
    "message_reviewed": False,
    "demo_logged_in": False,
    "demo_email": "",
    "demo_name": "",
}


for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# AUTHENTICATION / WELCOME GATE
# ============================================================

def has_real_oidc_config():
    try:
        return "auth" in st.secrets
    except Exception:
        return False


def show_login_screen():
    left, center, right = st.columns([0.65, 1.25, 0.8], gap="large")

    with center:
        st.markdown(
            """
            <div class="welcome-card">
                <div class="welcome-icon">♥</div>
                <div class="eyebrow">AI-Powered Discharge Documentation</div>
                <h1>CareNote <span style="color:#5d62e8">AI</span></h1>
                <div class="lead">Smarter documentation. Better patient care.</div>
                <h2 style="font-size:25px;margin-top:4px;">👋 Welcome!</h2>
                <p style="color:#7b8ba3;font-size:13px;">Sign in to continue to your CareNote AI workspace.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if has_real_oidc_config():
            if st.button("🌈  Continue with Google", type="primary", use_container_width=True):
                st.login()
        else:
            if st.button("🌈  Continue with Google", type="primary", use_container_width=True):
                st.session_state["demo_logged_in"] = True
                st.session_state["demo_name"] = "Demo User"
                st.session_state["demo_email"] = "demo@carenote.ai"
                st.rerun()
            st.caption("Google authentication will become real after OIDC is configured for the deployed app.")

        st.markdown("<div style='text-align:center;color:#9aa8bb;margin:15px 0;'>or</div>", unsafe_allow_html=True)

        with st.form("email_login_form"):
            email = st.text_input("Email", placeholder="you@example.com")
            name = st.text_input("Your name", placeholder="Your name")
            submitted = st.form_submit_button("✉  Continue with Email", use_container_width=True)
            if submitted:
                if "@" not in email or not name.strip():
                    st.warning("Please enter a name and a valid email address.")
                else:
                    st.session_state["demo_logged_in"] = True
                    st.session_state["demo_name"] = name.strip()
                    st.session_state["demo_email"] = email.strip().lower()
                    st.rerun()

        st.markdown(
            """
            <div class="security-box">🛡️ <b>Your information is secure and private.</b><br>
            CareNote AI is an AI documentation assistant. Use synthetic or anonymized data for this MVP.</div>
            """,
            unsafe_allow_html=True,
        )

    with right:
        st.markdown(
            """
            <div class="feature-panel">
                <div class="feature-photo">🩺</div>
                <div class="feature-item"><div class="feature-dot">⚡</div><div><b>AI-assisted summaries</b><span>Generate structured discharge documentation from supplied records.</span></div></div>
                <div class="feature-item"><div class="feature-dot">🔎</div><div><b>RAG-powered responses</b><span>Ground drafts in the hospital knowledge base.</span></div></div>
                <div class="feature-item"><div class="feature-dot">📄</div><div><b>Easy document upload</b><span>Work with PDF, CSV and TXT patient records.</span></div></div>
                <div class="feature-item"><div class="feature-dot">💬</div><div><b>Patient messaging</b><span>Create a patient-friendly message after professional review.</span></div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )


if has_real_oidc_config():
    try:
        logged_in = bool(st.user.is_logged_in)
    except Exception:
        logged_in = False
else:
    logged_in = bool(st.session_state.get("demo_logged_in", False))

if not logged_in:
    st.markdown(
        "<div class='topbar'><div class='topbar-title'>CareNote AI</div><div class='badge'>Secure workspace</div></div>",
        unsafe_allow_html=True,
    )
    show_login_screen()
    st.stop()

if has_real_oidc_config():
    CURRENT_USER_NAME = getattr(st.user, "name", None) or getattr(st.user, "email", None) or "CareNote User"
    CURRENT_USER_EMAIL = getattr(st.user, "email", "")
else:
    CURRENT_USER_NAME = st.session_state.get("demo_name") or "CareNote User"
    CURRENT_USER_EMAIL = st.session_state.get("demo_email", "")


# ============================================================
# BACKEND
# ============================================================

groq_api_key = os.getenv("GROQ_API_KEY")

if not groq_api_key:
    st.error("GROQ_API_KEY is not configured.")
    st.info(
        "For local/Streamlit Cloud deployment, add GROQ_API_KEY to your app secrets. "
        "Never place the API key directly inside app.py."
    )
    st.stop()

client = Groq(api_key=groq_api_key)


@st.cache_resource
def load_vectorstore():
    """Build the hospital knowledge-base vector store once per app instance."""
    kb_path = "data/knowledge_base/hospital_guidelines.txt"

    if not os.path.exists(kb_path):
        raise FileNotFoundError(
            f"Knowledge base not found: {kb_path}. "
            "Make sure data/knowledge_base/hospital_guidelines.txt exists."
        )

    loader = TextLoader(kb_path, encoding="utf-8")
    documents = loader.load()

    splitter = RecursiveCharacterTextSplitter(
        chunk_size=500,
        chunk_overlap=50,
    )
    chunks = splitter.split_documents(documents)

    embeddings = HuggingFaceEmbeddings(
        model_name="sentence-transformers/all-MiniLM-L6-v2"
    )

    return FAISS.from_documents(chunks, embeddings)


@st.cache_resource
def get_vectorstore():
    return load_vectorstore()


def extract_patient_data(uploaded_files):
    """Extract text from TXT, PDF and CSV uploads."""
    all_text = []

    for uploaded_file in uploaded_files:
        name = uploaded_file.name.lower()
        data = uploaded_file.getvalue()

        if name.endswith(".txt"):
            text = data.decode("utf-8", errors="ignore")

        elif name.endswith(".pdf"):
            reader = PdfReader(io.BytesIO(data))
            pages = []
            for page in reader.pages:
                pages.append(page.extract_text() or "")
            text = "\n".join(pages)

        elif name.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(data))
            text = df.to_string(index=False)

        else:
            raise ValueError(
                f"Unsupported file type: {uploaded_file.name}. "
                "Please upload PDF, CSV or TXT."
            )

        if text.strip():
            all_text.append(
                f"\n===== SOURCE DOCUMENT: {uploaded_file.name} =====\n{text.strip()}"
            )

    return "\n".join(all_text).strip()


def get_rag_sources(query, k=3):
    vectorstore = get_vectorstore()
    return vectorstore.similarity_search(query, k=k)


def _call_llm(system_prompt, user_prompt, temperature=0.1):
    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=temperature,
    )
    return response.choices[0].message.content.strip()


def generate_summary(patient_data, language="English"):
    """Generate a grounded discharge summary using patient data + hospital KB."""
    docs = get_rag_sources(patient_data, k=3)
    hospital_context = "\n\n".join(doc.page_content for doc in docs)

    language_instruction = """
Write the summary in English.
""" if language == "English" else """
Keep clinical/internal sections in English. Translate patient-facing instructions
and medication-related patient-facing text into clear Urdu where applicable.
Do not add any clinical information during translation.
"""

    system_prompt = """
You are a healthcare documentation assistant.

You are NOT a diagnostic or prescribing system.

Create a discharge-summary DRAFT using ONLY:
1. the supplied patient document content, and
2. the supplied hospital knowledge-base context.

Strict rules:
- Never invent patient facts, diagnoses, investigations, medications, dosages,
  appointments, dates, symptoms, findings, or treatment.
- Never infer a diagnosis that is not explicitly documented.
- Never change or create a medication or dosage.
- If a required item is absent, write "Not documented" or an equally clear
  statement rather than guessing.
- The hospital knowledge base is guidance for structure/documentation and does
  not prove medical correctness.
- Clearly label the result as an AI-generated draft requiring professional review.
- Use a clean, structured format with these sections:
  Patient Information
  Admission Information
  Diagnosis
  Investigations
  Treatment During Admission
  Condition at Discharge
  Medications
  Follow-up
  Patient Instructions
- Keep wording concise and professional.
"""

    user_prompt = f"""
PATIENT DOCUMENT CONTENT:
{patient_data}

HOSPITAL KNOWLEDGE-BASE CONTEXT:
{hospital_context}

LANGUAGE REQUIREMENT:
{language_instruction}

Generate the discharge-summary draft now.
"""

    return _call_llm(system_prompt, user_prompt, temperature=0.1)


def detect_missing_information(summary):
    missing_fields = []
    summary_lower = summary.lower()

    if "no discharge medications" in summary_lower:
        missing_fields.append("Discharge medications")

    if (
        "no specific follow-up" in summary_lower
        or (
            "follow-up" in summary_lower
            and "not documented" in summary_lower
        )
    ):
        missing_fields.append("Specific follow-up information")

    if (
        "specific additional patient instructions" in summary_lower
        and "not documented" in summary_lower
    ):
        missing_fields.append("Specific patient instructions")

    return missing_fields


def grounding_check(summary, patient_data):
    """Check whether obvious unsupported facts appear in the draft.

    This is a source-consistency check, not a medical-accuracy guarantee.
    """
    docs = get_rag_sources(patient_data, k=3)
    hospital_context = "\n\n".join(doc.page_content for doc in docs)

    system_prompt = """
You are a source-consistency reviewer.

Compare the candidate discharge summary against the supplied patient content
and hospital context.

Report only whether the summary contains information that is not supported by
the supplied sources.

Do NOT judge medical correctness.

Return exactly one of:
- "No unsupported information detected."
- "Potential unsupported information: <brief list>"

Be conservative: only flag claims that clearly cannot be traced to the sources.
"""

    prompt = f"""
PATIENT CONTENT:
{patient_data}

HOSPITAL CONTEXT:
{hospital_context}

CANDIDATE SUMMARY:
{summary}
"""

    return _call_llm(system_prompt, prompt, temperature=0.0)


def ask_document_question(question, patient_data, hospital_context=None):
    if hospital_context is None:
        docs = get_rag_sources(patient_data, k=3)
        hospital_context = "\n\n".join(doc.page_content for doc in docs)

    system_prompt = """
Answer the user's question using ONLY the supplied patient document and
hospital knowledge-base context.

If the answer is not present, say:
"Not documented in the available information."

Do not diagnose, prescribe, infer missing facts, or add medical advice.
Keep the answer concise.
"""

    prompt = f"""
PATIENT DOCUMENT:
{patient_data}

HOSPITAL CONTEXT:
{hospital_context}

QUESTION:
{question}
"""

    return _call_llm(system_prompt, prompt, temperature=0.0)


def regenerate_section(section_name, patient_data, language="English"):
    docs = get_rag_sources(patient_data, k=3)
    hospital_context = "\n\n".join(doc.page_content for doc in docs)

    system_prompt = """
You are a healthcare documentation assistant.

Regenerate ONLY the requested discharge-summary section.

Use only the patient document and hospital context.
Do not invent, infer, diagnose, prescribe, change medications, or add dates.
If information is absent, say "Not documented."

Return only the replacement section text.
"""

    prompt = f"""
REQUESTED SECTION:
{section_name}

PATIENT DOCUMENT:
{patient_data}

HOSPITAL CONTEXT:
{hospital_context}

LANGUAGE:
{language}
"""

    return _call_llm(system_prompt, prompt, temperature=0.1)


# ============================================================
# EXPORTS
# ============================================================

def generate_pdf(text):
    """Create PDF from the exact edited summary text."""
    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=48,
        leftMargin=48,
        topMargin=48,
        bottomMargin=48,
        title="AI Discharge Summary",
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TitleCustom",
        parent=styles["Title"],
        fontSize=20,
        leading=24,
        alignment=TA_CENTER,
        spaceAfter=18,
    )

    body_style = ParagraphStyle(
        "BodyCustom",
        parent=styles["BodyText"],
        fontSize=10.5,
        leading=15,
        spaceAfter=7,
    )

    story = [
        Paragraph("AI Discharge Summary", title_style),
    ]

    for raw_line in text.splitlines():
        line = raw_line.strip()

        if not line:
            story.append(Spacer(1, 5))
            continue

        safe = html.escape(line)
        safe = safe.replace("**", "")

        if re.match(r"^\d+\.\s+\*\*.*\*\*$", line):
            safe = re.sub(r"\*\*", "", line)
            story.append(Paragraph(f"<b>{html.escape(safe)}</b>", body_style))
        elif line.startswith("**") and line.endswith("**"):
            safe = line.strip("*")
            story.append(Paragraph(f"<b>{html.escape(safe)}</b>", body_style))
        elif line.startswith("- "):
            story.append(Paragraph(f"• {html.escape(line[2:])}", body_style))
        else:
            story.append(Paragraph(safe, body_style))

    story.append(Spacer(1, 15))
    story.append(
        Paragraph(
            f"<b>Disclaimer:</b> {html.escape(DISCLAIMER)}",
            ParagraphStyle(
                "Disclaimer",
                parent=body_style,
                fontSize=8.5,
                leading=12,
                textColor=colors.grey,
            ),
        )
    )

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def generate_docx(text):
    """Create DOCX from the exact edited summary text."""
    document = Document()

    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    title = document.add_heading("AI Discharge Summary", level=0)
    title.alignment = 1

    for raw_line in text.splitlines():
        line = raw_line.strip()

        if not line:
            document.add_paragraph("")
            continue

        clean = line.replace("**", "")

        if re.match(r"^\d+\.\s+", clean):
            p = document.add_paragraph()
            run = p.add_run(clean)
            run.bold = True
            run.font.size = Pt(12)
        elif clean.startswith("- "):
            p = document.add_paragraph(style="List Bullet")
            p.add_run(clean[2:])
        else:
            p = document.add_paragraph(clean)
            for run in p.runs:
                run.font.size = Pt(10.5)

    document.add_paragraph("")
    p = document.add_paragraph()
    run = p.add_run("Disclaimer: " + DISCLAIMER)
    run.italic = True
    run.font.size = Pt(8.5)

    buffer = io.BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.markdown(
        f"""
        <div class="brand-wrap">
            <span class="brand-mark">♥</span><span class="brand-name">CareNote <span class="brand-ai">AI</span></span>
            <div class="brand-sub">AI-Powered Discharge Summaries</div>
        </div>
        <div class="sidebar-label">Workspace</div>
        """,
        unsafe_allow_html=True,
    )

    page = st.radio(
        "Navigation",
        [
            "🏠 Home",
            "📄 Patient Documents",
            "📋 Discharge Summary",
            "💬 Document Q&A",
            "💌 Patient Message",
            "🔎 RAG Sources",
            "ℹ️ About",
        ],
        label_visibility="collapsed",
    )

    st.markdown('<div class="sidebar-label">System</div>', unsafe_allow_html=True)
    st.markdown('<span class="badge">● RAG ready</span>', unsafe_allow_html=True)
    st.caption(f"Model: {MODEL_NAME}")
    st.caption("FAISS · Hugging Face embeddings")

    st.markdown('<div class="sidebar-label">Account</div>', unsafe_allow_html=True)
    st.caption(f"👤 {CURRENT_USER_NAME}")

    if has_real_oidc_config():
        if st.button("↪ Log out", use_container_width=True):
            st.logout()
    else:
        if st.button("↪ Log out", use_container_width=True):
            st.session_state["demo_logged_in"] = False
            st.session_state["demo_email"] = ""
            st.session_state["demo_name"] = ""
            st.rerun()

    st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)
    if st.button("＋ Start new case", use_container_width=True):
        keep = {"demo_logged_in", "demo_email", "demo_name"}
        for key, value in DEFAULTS.items():
            if key not in keep:
                st.session_state[key] = value
        st.rerun()


# ============================================================
# TOP HEADER
# ============================================================

st.markdown(
    f"""
    <div class="topbar">
        <div class="topbar-title">Healthcare documentation workspace</div>
        <div class="user-chip"><span class="avatar">👤</span> Welcome, {html.escape(str(CURRENT_USER_NAME))}</div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HOME / DASHBOARD
# ============================================================

if page == "🏠 Home":
    st.markdown(
        f"""
        <div class="hero">
            <div class="eyebrow">CareNote AI · Your workspace</div>
            <h1>Welcome back, {html.escape(str(CURRENT_USER_NAME))} 👋</h1>
            <p>Turn source records into structured, reviewable discharge documentation with retrieval-augmented generation.</p>
            <div class="hero-badges"><span class="badge">RAG grounded</span><span class="badge">Human review</span><span class="badge">PDF + Word</span><span class="badge">English + Urdu</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1,c2,c3,c4=st.columns(4)
    with c1: st.markdown(f"<div class='metric'><div class='label'>Documents</div><div class='value'>{len(st.session_state['uploaded_names'])}</div></div>", unsafe_allow_html=True)
    with c2: st.markdown(f"<div class='metric'><div class='label'>Summary</div><div class='value'>{'Ready' if st.session_state['edited_summary'] else 'Not started'}</div></div>", unsafe_allow_html=True)
    with c3: st.markdown(f"<div class='metric'><div class='label'>RAG sources</div><div class='value'>{len(st.session_state['rag_sources'])}</div></div>", unsafe_allow_html=True)
    with c4: st.markdown(f"<div class='metric'><div class='label'>Message</div><div class='value'>{'Ready' if st.session_state['patient_message'] else 'Not created'}</div></div>", unsafe_allow_html=True)

    st.write("")
    left,right=st.columns([1.45,1],gap="large")
    with left:
        st.markdown("<div class='card'><div class='card-title'>How CareNote AI works</div><div class='card-sub'>A simple human-in-the-loop workflow</div><p>📄 <b>Upload records</b> → 🔎 <b>Retrieve hospital guidance</b> → ✨ <b>Generate draft</b> → 👩‍⚕️ <b>Review & edit</b> → 📥 <b>Export</b> → 💌 <b>Prepare patient message</b></p></div>", unsafe_allow_html=True)
    with right:
        st.markdown("<div class='card'><div class='card-title'>Ready to start?</div><div class='card-sub'>Use synthetic or anonymized data for this MVP.</div></div>", unsafe_allow_html=True)
        if st.button("＋ Create New Discharge Summary", type="primary", use_container_width=True):
            st.session_state["_go_documents"] = True
            st.rerun()

    if st.session_state.get("_go_documents"):
        st.session_state["_go_documents"] = False
        st.info("Choose **Patient Documents** from the sidebar to upload a record.")

    st.markdown("<div class='warning'><b>Clinical safety:</b> Every AI-generated summary must be reviewed and approved by a qualified healthcare professional. This MVP does not diagnose, prescribe, or change treatment.</div>", unsafe_allow_html=True)


# ============================================================
# PAGE 1 — PATIENT DOCUMENTS
# ============================================================

elif page == "📄 Patient Documents":

    st.markdown("<div class='eyebrow'>Patient workspace</div><h2 style='margin-top:4px'>Patient Documents</h2><p style='color:#7b8ba3'>Upload source records and create a grounded discharge-summary draft.</p>", unsafe_allow_html=True)

    left, right = st.columns([1.65, 1], gap="large")

    with left:
        uploaded_files = st.file_uploader(
            "Upload patient documents",
            type=["pdf", "csv", "txt"],
            accept_multiple_files=True,
            help="For the hackathon MVP, use synthetic or anonymized text-based documents.",
        )

        if uploaded_files:
            st.success(f"{len(uploaded_files)} document(s) ready.")

            names = [f.name for f in uploaded_files]
            st.session_state["uploaded_names"] = names

            for name in names:
                st.markdown(
                    f"""
                    <div class="section-card">
                        <div class="section-title">📄 {html.escape(name)}</div>
                        <div class="section-help">Ready for text extraction</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    with right:
        st.markdown(
            """
            <div class="section-card">
                <div class="section-title">How it works</div>
                <div class="section-help">
                    Your uploaded content remains the primary source. RAG adds
                    hospital documentation guidance before generation.
                </div>
                <p><b>1.</b> Extract patient information</p>
                <p><b>2.</b> Retrieve relevant hospital guidance</p>
                <p><b>3.</b> Generate a grounded draft</p>
                <p><b>4.</b> Review and edit</p>
                <p><b>5.</b> Export PDF + Word</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        language = st.selectbox(
            "Patient-facing language",
            ["English", "Urdu"],
            index=0 if st.session_state["language"] == "English" else 1,
            help="Clinical/internal sections remain in English for the MVP; patient-facing content can be requested in Urdu.",
        )
        st.session_state["language"] = language

    if uploaded_files:
        if st.button("✨ Extract & Generate Summary", type="primary", use_container_width=True):
            try:
                with st.spinner("Extracting documents and retrieving hospital guidance..."):
                    patient_data = extract_patient_data(uploaded_files)

                    if not patient_data.strip():
                        st.error("No readable text was found in the uploaded documents.")
                        st.stop()

                    summary = generate_summary(
                        patient_data,
                        st.session_state["language"],
                    )

                st.session_state["patient_data"] = patient_data
                st.session_state["discharge_summary"] = summary
                st.session_state["edited_summary"] = summary
                st.session_state["missing_fields"] = detect_missing_information(summary)
                st.session_state["grounding_result"] = ""
                st.session_state["rag_sources"] = get_rag_sources(patient_data, k=3)

                st.success("Summary generated successfully.")
                st.info("Open **Review & Export** from the sidebar to review and edit the draft.")

            except Exception as e:
                st.error(f"Generation failed: {e}")
                st.caption("Check your API key, knowledge-base file, uploaded document, and internet connection.")

    if st.session_state["patient_data"]:
        with st.expander("View extracted patient information"):
            st.text(st.session_state["patient_data"])


# ============================================================
# PAGE 2 — REVIEW & EXPORT
# ============================================================

elif page == "📋 Discharge Summary":

    st.subheader("Review & finalize")
    st.caption("Step 2 of 3 · Human review before export")

    if not st.session_state["discharge_summary"]:
        st.info("No summary is available yet. Create a new summary first.")
    else:
        summary_col, tools_col = st.columns([1.75, 1], gap="large")

        with summary_col:
            st.markdown(
                """
                <div class="section-card">
                    <div class="section-title">Editable discharge summary</div>
                    <div class="section-help">
                        Make any required corrections before exporting. Downloads use
                        the exact text in this editor.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            edited = st.text_area(
                "Final summary",
                value=st.session_state["edited_summary"],
                height=650,
                label_visibility="collapsed",
            )
            st.session_state["edited_summary"] = edited

        with tools_col:
            st.markdown("### Review checks")

            missing = detect_missing_information(edited)

            if missing:
                st.warning("Some information is explicitly not documented.")
                for item in missing:
                    st.markdown(f"• {item}")
            else:
                st.success("No predefined missing-information flags found.")

            if st.button("🔎 Run source consistency check", use_container_width=True):
                with st.spinner("Comparing draft with source content..."):
                    result = grounding_check(
                        edited,
                        st.session_state["patient_data"],
                    )
                st.session_state["grounding_result"] = result

            if st.session_state["grounding_result"]:
                result = st.session_state["grounding_result"]
                if result.startswith("No unsupported"):
                    st.success(result)
                else:
                    st.warning(result)

            st.markdown("### Regenerate one section")

            section = st.selectbox(
                "Section",
                [
                    "Patient Information",
                    "Admission Information",
                    "Diagnosis",
                    "Investigations",
                    "Treatment During Admission",
                    "Condition at Discharge",
                    "Medications",
                    "Follow-up",
                    "Patient Instructions",
                ],
            )

            if st.button("↻ Regenerate selected section", use_container_width=True):
                with st.spinner(f"Regenerating {section}..."):
                    replacement = regenerate_section(
                        section,
                        st.session_state["patient_data"],
                        st.session_state["language"],
                    )
                st.session_state["section_preview"] = replacement
                st.rerun()

            if st.session_state.get("section_preview"):
                st.markdown("**Replacement preview**")
                st.code(st.session_state["section_preview"], language="text")

                if st.button("Use this replacement", use_container_width=True):
                    st.session_state["edited_summary"] += (
                        "\n\n" + st.session_state["section_preview"]
                    )
                    st.session_state["section_preview"] = ""
                    st.rerun()

            st.markdown("### Export")

            final_text = st.session_state["edited_summary"].strip()

            if final_text:
                pdf_bytes = generate_pdf(final_text)
                docx_bytes = generate_docx(final_text)

                st.download_button(
                    "⬇ Download PDF",
                    data=pdf_bytes,
                    file_name="discharge_summary_final.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

                st.download_button(
                    "⬇ Download Word (.docx)",
                    data=docx_bytes,
                    file_name="discharge_summary_final.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            else:
                st.info("Enter or generate a summary before exporting.")

        st.markdown(
            """
            <div class="footer-note">
                <b>Human review required:</b> This application produces documentation
                drafts from supplied sources. It does not independently verify medical
                correctness or replace professional clinical judgment.
            </div>
            """,
            unsafe_allow_html=True,
        )


# ============================================================
# PAGE 3 — DOCUMENT Q&A
# ============================================================

elif page == "💬 Document Q&A":

    st.subheader("Ask the patient documents")
    st.caption("Step 3 · Source-grounded question answering")

    if not st.session_state["patient_data"]:
        st.info("Upload and process a patient document first.")
    else:
        st.markdown(
            """
            <div class="info-banner">
                Ask questions about information contained in the uploaded patient
                documents. The assistant is instructed not to invent missing facts.
            </div>
            """,
            unsafe_allow_html=True,
        )

        question = st.text_input(
            "Your question",
            placeholder="Example: Were any discharge medications documented?",
        )

        if st.button("Ask question", type="primary"):
            if not question.strip():
                st.warning("Please enter a question.")
            else:
                with st.spinner("Searching sources and answering..."):
                    answer = ask_document_question(
                        question,
                        st.session_state["patient_data"],
                    )
                st.session_state["qa_question"] = question
                st.session_state["qa_answer"] = answer

        if st.session_state["qa_answer"]:
            st.markdown("### Answer")
            st.markdown(
                f"""
                <div class="section-card">
                    {html.escape(st.session_state["qa_answer"]).replace(chr(10), "<br>")}
                </div>
                """,
                unsafe_allow_html=True,
            )

        with st.expander("View source patient information"):
            st.text(st.session_state["patient_data"])


# ============================================================
# PATIENT MESSAGE
# ============================================================

elif page == "💌 Patient Message":
    st.markdown("<div class='eyebrow'>Patient communication</div><h2 style='margin-top:4px'>Patient Message</h2><p style='color:#7b8ba3'>Create a clear, patient-friendly communication from the reviewed discharge summary.</p>", unsafe_allow_html=True)

    if not st.session_state["edited_summary"]:
        st.info("Create and review a discharge summary first.")
    else:
        st.markdown("<div class='notice'><b>Important:</b> The message is a communication draft. A qualified professional should review it before any real-world communication. The MVP does not automatically send messages.</div>", unsafe_allow_html=True)
        st.write("")
        tone = st.selectbox("Message style", ["Clear and friendly", "Formal and concise", "Simple language"])
        if st.button("✨ Create Patient Message", type="primary", use_container_width=True):
            system = """You are a healthcare documentation communication assistant. Convert the reviewed discharge summary into a short patient-facing message. Use ONLY information explicitly present in the summary. Do not diagnose, prescribe, invent appointments, add dosages, or introduce new clinical advice. Mention that the message is a draft for professional review. Keep it clear and respectful."""
            prompt = f"Reviewed discharge summary:\n{st.session_state['edited_summary']}\n\nStyle: {tone}\nLanguage: {st.session_state['language']}\nCreate the message."
            with st.spinner("Preparing patient message..."):
                st.session_state["patient_message"] = _call_llm(system, prompt, temperature=0.1)

        if st.session_state["patient_message"]:
            st.markdown("### Message preview")
            message = st.text_area("Final patient message", value=st.session_state["patient_message"], height=300, key="patient_message_editor")
            st.session_state["patient_message"] = message

            a,b=st.columns(2)
            with a:
                if st.button("✓ Mark as reviewed", use_container_width=True):
                    st.session_state["message_reviewed"] = True
                    st.success("Marked as reviewed for the demo.")
            with b:
                if st.button("📨 Demo send", use_container_width=True):
                    st.success("Demo send completed. No real patient message was sent.")

            if st.session_state.get("message_reviewed"):
                st.markdown("<div class='success-box'><b>Reviewed.</b> This prototype can now demonstrate a send action without contacting a real patient.</div>", unsafe_allow_html=True)


# ============================================================
# ABOUT
# ============================================================

elif page == "ℹ️ About":
    st.markdown("<div class='eyebrow'>CareNote AI</div><h2 style='margin-top:4px'>About the MVP</h2>", unsafe_allow_html=True)
    st.markdown("""
    <div class='card'><div class='card-title'>What CareNote AI does</div><div class='card-sub'>RAG-assisted clinical documentation</div>
    <p>CareNote AI extracts information from supplied patient records, retrieves relevant hospital guidance, and generates an editable discharge-summary draft.</p>
    <p><b>Technology:</b> Streamlit · Groq · GPT-OSS 20B · FAISS · Hugging Face embeddings · LangChain</p></div>
    <div class='warning'><b>Safety:</b> This is a documentation assistant, not a diagnostic or prescribing system. Use synthetic/anonymized data for the MVP. Every generated document requires qualified professional review.</div>
    """, unsafe_allow_html=True)


# ============================================================
# PAGE 4 — RAG SOURCES
# ============================================================

elif page == "🔎 RAG Sources":

    st.subheader("RAG transparency")
    st.caption("See the hospital knowledge-base context retrieved for this case.")

    if not st.session_state["patient_data"]:
        st.info("Generate a summary first to populate the retrieved sources.")
    else:
        if not st.session_state["rag_sources"]:
            st.session_state["rag_sources"] = get_rag_sources(
                st.session_state["patient_data"],
                k=3,
            )

        st.markdown(
            """
            <div class="info-banner">
                These retrieved chunks are used as grounding context for generation.
                Retrieval does not guarantee clinical correctness.
            </div>
            """,
            unsafe_allow_html=True,
        )

        for i, doc in enumerate(st.session_state["rag_sources"], start=1):
            st.markdown(f"### Source {i}")
            st.markdown(
                f"""
                <div class="section-card">
                    {html.escape(doc.page_content).replace(chr(10), "<br>")}
                </div>
                """,
                unsafe_allow_html=True,
            )


# ============================================================
# GLOBAL FOOTER
# ============================================================

st.markdown(
    f"""
    <div class="footer-note">
        <b>CareNote AI</b> · RAG-assisted discharge documentation · MVP<br>
        {html.escape(DISCLAIMER)}
    </div>
    """,
    unsafe_allow_html=True,
)
