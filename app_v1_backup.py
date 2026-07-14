"""
================================================================================
 ResumeBoost AI Pro
 Land More Interviews. Build a Stronger Career.

 A fully offline, rule-based AI Career Success Platform.
 Built for Next Byte Hacks V3.

 No internet. No external APIs. 100% local Python + Streamlit logic.
================================================================================
"""

import re
import io
import random
import textwrap
from datetime import datetime

import streamlit as st
import pandas as pd

# Optional parsers - handled gracefully if unavailable
try:
    import pdfplumber
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

try:
    import docx
    DOCX_SUPPORT = True
except ImportError:
    DOCX_SUPPORT = False

try:
    from fpdf import FPDF
    FPDF_SUPPORT = True
except ImportError:
    FPDF_SUPPORT = False


# ==============================================================================
# SECTION 1: PAGE CONFIGURATION
# ==============================================================================

st.set_page_config(
    page_title="ResumeBoost AI Pro | Next Byte Hacks V3",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==============================================================================
# SECTION 2: GLOBAL STYLING (DARK THEME + BLUE/PURPLE GRADIENT SAAS LOOK)
# ==============================================================================

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Poppins:wght@600;700;800&display=swap');

html, body, [class*="css"]  {
    font-family: 'Inter', sans-serif;
}

/* App background */
.stApp {
    background: radial-gradient(circle at 10% 0%, #1b1035 0%, #0d0b1e 45%, #090815 100%);
    color: #EAEAF5;
}

/* Hide default Streamlit chrome */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

/* Hero container */
.hero-wrap {
    background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 45%, #C026D3 100%);
    padding: 3rem 2.5rem;
    border-radius: 24px;
    margin-bottom: 2rem;
    box-shadow: 0 20px 60px rgba(124, 58, 237, 0.35);
    position: relative;
    overflow: hidden;
}
.hero-title {
    font-family: 'Poppins', sans-serif;
    font-size: 3rem;
    font-weight: 800;
    color: white;
    margin-bottom: 0.4rem;
    letter-spacing: -1px;
}
.hero-sub {
    font-size: 1.25rem;
    color: rgba(255,255,255,0.92);
    font-weight: 500;
    margin-bottom: 1rem;
}
.hero-badge {
    display: inline-block;
    background: rgba(255,255,255,0.18);
    border: 1px solid rgba(255,255,255,0.35);
    color: white;
    padding: 0.35rem 0.9rem;
    border-radius: 999px;
    font-size: 0.8rem;
    font-weight: 600;
    margin-bottom: 1rem;
    backdrop-filter: blur(6px);
}

/* Generic glass card */
.gb-card {
    background: linear-gradient(160deg, rgba(255,255,255,0.06), rgba(255,255,255,0.02));
    border: 1px solid rgba(255,255,255,0.09);
    border-radius: 18px;
    padding: 1.4rem 1.5rem;
    margin-bottom: 1rem;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}
.gb-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 14px 30px rgba(124, 58, 237, 0.25);
    border: 1px solid rgba(167, 139, 250, 0.5);
}
.gb-feature-icon {
    font-size: 2rem;
    margin-bottom: 0.5rem;
}
.gb-feature-title {
    font-weight: 700;
    font-size: 1.05rem;
    color: #F5F3FF;
    margin-bottom: 0.3rem;
}
.gb-feature-desc {
    font-size: 0.88rem;
    color: #C4C1DD;
    line-height: 1.4;
}

/* Section headers */
.section-title {
    font-family: 'Poppins', sans-serif;
    font-size: 1.6rem;
    font-weight: 700;
    color: #F5F3FF;
    margin: 1.6rem 0 0.8rem 0;
    border-left: 5px solid #A78BFA;
    padding-left: 0.7rem;
}

/* Score cards */
.score-card {
    background: linear-gradient(160deg, rgba(124,58,237,0.18), rgba(79,70,229,0.08));
    border: 1px solid rgba(167,139,250,0.35);
    border-radius: 16px;
    padding: 1.1rem 1rem;
    text-align: center;
}
.score-value {
    font-size: 2.1rem;
    font-weight: 800;
    font-family: 'Poppins', sans-serif;
    background: linear-gradient(90deg, #A78BFA, #F472B6);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}
.score-label {
    font-size: 0.82rem;
    color: #C4C1DD;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

/* Badges */
.badge-good { background:#064e3b; color:#6ee7b7; padding:3px 10px; border-radius:8px; font-size:0.78rem; font-weight:600;}
.badge-warn { background:#5f3a00; color:#fbbf24; padding:3px 10px; border-radius:8px; font-size:0.78rem; font-weight:600;}
.badge-bad  { background:#5b1a1a; color:#fca5a5; padding:3px 10px; border-radius:8px; font-size:0.78rem; font-weight:600;}

/* Chips for keywords */
.chip-match {
    display:inline-block; background:rgba(16,185,129,0.15); color:#6ee7b7;
    border:1px solid rgba(16,185,129,0.4); padding:4px 12px; border-radius:999px;
    font-size:0.82rem; margin:3px; font-weight:600;
}
.chip-missing {
    display:inline-block; background:rgba(239,68,68,0.15); color:#fca5a5;
    border:1px solid rgba(239,68,68,0.4); padding:4px 12px; border-radius:999px;
    font-size:0.82rem; margin:3px; font-weight:600;
}

/* Custom progress bar */
.progress-outer {
    background: rgba(255,255,255,0.08);
    border-radius: 999px;
    height: 14px;
    width: 100%;
    overflow: hidden;
    margin: 6px 0 14px 0;
}
.progress-inner {
    height: 100%;
    border-radius: 999px;
    background: linear-gradient(90deg, #6366F1, #C026D3);
    transition: width 1.2s ease-in-out;
}

/* Alert boxes for scam detector */
.alert-danger {
    background: rgba(239,68,68,0.12);
    border-left: 4px solid #ef4444;
    padding: 0.8rem 1rem;
    border-radius: 10px;
    margin-bottom: 0.6rem;
    color: #fecaca;
}
.alert-safe {
    background: rgba(16,185,129,0.12);
    border-left: 4px solid #10b981;
    padding: 0.8rem 1rem;
    border-radius: 10px;
    margin-bottom: 0.6rem;
    color: #a7f3d0;
}

/* Footer */
.gb-footer {
    text-align: center;
    color: #8b87ab;
    font-size: 0.85rem;
    padding: 2rem 0 1rem 0;
    border-top: 1px solid rgba(255,255,255,0.08);
    margin-top: 3rem;
}

/* Sidebar styling */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #150f2e 0%, #0d0b1e 100%);
    border-right: 1px solid rgba(255,255,255,0.06);
}

/* Buttons */
.stButton>button {
    background: linear-gradient(90deg, #6366F1, #C026D3);
    color: white;
    border: none;
    border-radius: 12px;
    padding: 0.55rem 1.4rem;
    font-weight: 700;
    transition: all 0.2s ease;
}
.stButton>button:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 20px rgba(124,58,237,0.4);
    color: white;
}

/* Text areas / inputs */
.stTextArea textarea, .stTextInput input {
    background: rgba(255,255,255,0.04) !important;
    color: #EAEAF5 !important;
    border-radius: 12px !important;
    border: 1px solid rgba(255,255,255,0.12) !important;
}

/* Bullet rewrite box */
.rewrite-before {
    background: rgba(239,68,68,0.08);
    border-left: 3px solid #ef4444;
    padding: 0.7rem 1rem;
    border-radius: 8px;
    margin-bottom: 0.3rem;
    font-size: 0.9rem;
    color: #fca5a5;
}
.rewrite-after {
    background: rgba(16,185,129,0.08);
    border-left: 3px solid #10b981;
    padding: 0.7rem 1rem;
    border-radius: 8px;
    margin-bottom: 0.8rem;
    font-size: 0.9rem;
    color: #a7f3d0;
}

/* Roadmap week card */
.week-card {
    background: linear-gradient(160deg, rgba(99,102,241,0.12), rgba(192,38,211,0.06));
    border: 1px solid rgba(167,139,250,0.3);
    border-radius: 16px;
    padding: 1.2rem 1.4rem;
    margin-bottom: 1rem;
}
.week-title {
    font-weight: 800;
    font-size: 1.1rem;
    color: #E9D5FF;
    margin-bottom: 0.5rem;
}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ==============================================================================
# SECTION 3: TEXT EXTRACTION HELPERS (PDF / DOCX / TXT)
# ==============================================================================

def extract_text_from_pdf(file_obj) -> str:
    """Extract raw text from an uploaded PDF file object."""
    if not PDF_SUPPORT:
        return ""
    text_chunks = []
    try:
        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text_chunks.append(page_text)
    except Exception:
        return ""
    return "\n".join(text_chunks)


def extract_text_from_docx(file_obj) -> str:
    """Extract raw text from an uploaded DOCX file object."""
    if not DOCX_SUPPORT:
        return ""
    try:
        document = docx.Document(file_obj)
        paragraphs = [p.text for p in document.paragraphs if p.text.strip()]
        return "\n".join(paragraphs)
    except Exception:
        return ""


def extract_resume_text(uploaded_file) -> str:
    """Route an uploaded file to the correct extractor based on extension."""
    if uploaded_file is None:
        return ""
    filename = uploaded_file.name.lower()
    try:
        if filename.endswith(".pdf"):
            return extract_text_from_pdf(uploaded_file)
        elif filename.endswith(".docx"):
            return extract_text_from_docx(uploaded_file)
        elif filename.endswith(".txt"):
            return uploaded_file.read().decode("utf-8", errors="ignore")
        else:
            return ""
    except Exception:
        return ""


# ==============================================================================
# SECTION 4: RESUME ANALYSIS ENGINE (RULE-BASED NLP HEURISTICS)
# ==============================================================================

ACTION_VERBS = {
    "achieved", "accelerated", "administered", "advised", "analyzed", "architected",
    "automated", "built", "boosted", "championed", "coordinated", "created",
    "delivered", "designed", "developed", "directed", "drove", "engineered",
    "established", "executed", "expanded", "generated", "implemented", "improved",
    "increased", "initiated", "launched", "led", "managed", "mentored", "negotiated",
    "optimized", "orchestrated", "organized", "overhauled", "pioneered", "produced",
    "reduced", "redesigned", "resolved", "restructured", "scaled", "shipped",
    "simplified", "spearheaded", "streamlined", "strengthened", "supervised",
    "transformed", "upgraded", "won",
}

WEAK_VERBS = {
    "worked", "helped", "did", "made", "handled", "was responsible for",
    "responsible for", "involved in", "assisted", "participated", "tried",
    "dealt with", "used", "did work", "took part", "got", "went",
}

FILLER_WORDS = {
    "very", "really", "just", "basically", "actually", "literally", "stuff",
    "things", "a lot", "etc", "various", "several", "some", "kind of",
    "sort of", "quite", "highly motivated", "hard worker", "team player",
    "detail oriented", "go-getter", "think outside the box", "synergy",
}

PASSIVE_PATTERNS = [
    r"\bwas\s+\w+ed\b", r"\bwere\s+\w+ed\b", r"\bis\s+\w+ed\b",
    r"\bbeen\s+\w+ed\b", r"\bbeing\s+\w+ed\b", r"\bwas\s+\w+en\b",
    r"\bwere\s+\w+en\b",
]

BULLET_SPLIT_PATTERN = re.compile(r"[\n•\-\u2022]+")


def clean_lines(text: str):
    """Split resume text into non-empty stripped lines."""
    return [ln.strip() for ln in text.split("\n") if ln.strip()]


def count_action_verbs(text: str) -> int:
    words = re.findall(r"\b[a-zA-Z]+\b", text.lower())
    return sum(1 for w in words if w in ACTION_VERBS)


def count_weak_verbs(text: str) -> int:
    lower_text = text.lower()
    return sum(lower_text.count(w) for w in WEAK_VERBS)


def count_filler_words(text: str) -> int:
    lower_text = text.lower()
    return sum(lower_text.count(w) for w in FILLER_WORDS)


def count_passive_voice(text: str) -> int:
    lower_text = text.lower()
    total = 0
    for pattern in PASSIVE_PATTERNS:
        total += len(re.findall(pattern, lower_text))
    return total


def count_measurable_achievements(text: str) -> int:
    """Detect numbers, percentages, dollar amounts, and metric-style claims."""
    patterns = [
        r"\d+%", r"\$\d+", r"\d+x\b", r"\b\d{2,}\b",
        r"\d+\s?(hours|days|weeks|months|years|users|clients|customers|projects|team members)",
    ]
    total = 0
    for pattern in patterns:
        total += len(re.findall(pattern, text.lower()))
    return total


def estimate_readability(text: str) -> float:
    """
    Approximate a Flesch-Reading-Ease-like score using simple heuristics
    (word length + sentence length), fully offline, no external corpora.
    """
    sentences = re.split(r"[.!?]+", text)
    sentences = [s for s in sentences if s.strip()]
    words = re.findall(r"\b[a-zA-Z]+\b", text)

    if not sentences or not words:
        return 50.0

    avg_sentence_len = len(words) / max(len(sentences), 1)
    avg_word_len = sum(len(w) for w in words) / max(len(words), 1)

    # Simplified readability formula: shorter sentences/words = higher score
    score = 206.835 - (1.015 * avg_sentence_len) - (84.6 * (avg_word_len / 10))
    return max(0, min(100, round(score, 1)))


def detect_resume_sections(text: str) -> dict:
    """Detect presence of common resume sections."""
    lower = text.lower()
    sections = {
        "Contact Info": bool(re.search(r"@[\w\.-]+\.\w+|\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}", text)),
        "Summary/Objective": bool(re.search(r"\b(summary|objective|profile)\b", lower)),
        "Experience": bool(re.search(r"\b(experience|employment|work history)\b", lower)),
        "Education": bool(re.search(r"\b(education|degree|university|college|bachelor|master)\b", lower)),
        "Skills": bool(re.search(r"\b(skills|technologies|proficienc)\b", lower)),
        "Projects": bool(re.search(r"\b(projects?)\b", lower)),
    }
    return sections


def compute_resume_scores(text: str) -> dict:
    """
    Master scoring function. Produces a dictionary of scores (0-100) covering
    impact, ATS-friendliness, readability, and overall career readiness.
    """
    word_count = len(re.findall(r"\b[a-zA-Z]+\b", text))
    action_verbs = count_action_verbs(text)
    weak_verbs = count_weak_verbs(text)
    filler = count_filler_words(text)
    passive = count_passive_voice(text)
    measurable = count_measurable_achievements(text)
    readability = estimate_readability(text)
    sections = detect_resume_sections(text)
    section_score = (sum(sections.values()) / len(sections)) * 100

    # --- Impact Score: rewards action verbs & measurable achievements ---
    impact_score = 50
    impact_score += min(25, action_verbs * 2.5)
    impact_score += min(20, measurable * 2)
    impact_score -= min(20, weak_verbs * 3)
    impact_score -= min(15, passive * 3)
    impact_score = max(0, min(100, round(impact_score)))

    # --- ATS Score: rewards structure, keywords, standard sections ---
    ats_score = 40
    ats_score += section_score * 0.35
    ats_score += min(15, action_verbs * 1.2)
    ats_score -= min(15, filler * 2)
    length_penalty = 0
    if word_count < 150:
        length_penalty = 15
    elif word_count > 1000:
        length_penalty = 10
    ats_score -= length_penalty
    ats_score = max(0, min(100, round(ats_score)))

    # --- Readability normalized to 0-100 already ---
    readability_score = round(readability)

    # --- Overall Score: weighted composite ---
    overall_score = round(
        (impact_score * 0.35) + (ats_score * 0.35) + (readability_score * 0.15) + (section_score * 0.15)
    )
    overall_score = max(0, min(100, overall_score))

    # --- Career Readiness: composite of everything + resume length sanity ---
    readiness = round((overall_score * 0.6) + (section_score * 0.4))
    readiness = max(0, min(100, readiness))

    return {
        "word_count": word_count,
        "action_verbs": action_verbs,
        "weak_verbs": weak_verbs,
        "filler_words": filler,
        "passive_voice": passive,
        "measurable_achievements": measurable,
        "readability": readability_score,
        "sections": sections,
        "section_score": round(section_score),
        "impact_score": impact_score,
        "ats_score": ats_score,
        "overall_score": overall_score,
        "career_readiness": readiness,
    }


def score_badge(score: int) -> str:
    """Return an HTML badge class label based on score thresholds."""
    if score >= 75:
        return '<span class="badge-good">Strong</span>'
    elif score >= 50:
        return '<span class="badge-warn">Needs Work</span>'
    else:
        return '<span class="badge-bad">Weak</span>'


# ==============================================================================
# SECTION 5: ATS JOB MATCH ENGINE
# ==============================================================================

COMMON_STOPWORDS = {
    "the", "and", "a", "an", "to", "of", "in", "for", "with", "on", "is", "are",
    "as", "be", "this", "that", "will", "you", "your", "we", "our", "or", "at",
    "by", "from", "have", "has", "it", "its", "their", "they", "who", "which",
    "not", "but", "can", "must", "should", "would", "into", "about", "such",
    "including", "etc", "job", "role", "work", "working", "years", "year",
    "experience", "ability", "skills", "team", "company", "position",
}


def extract_keywords(text: str, top_n: int = 40):
    """Extract meaningful keywords/phrases from a job description."""
    words = re.findall(r"\b[a-zA-Z][a-zA-Z0-9\+\#\.]{2,}\b", text.lower())
    filtered = [w for w in words if w not in COMMON_STOPWORDS and len(w) > 2]

    freq = {}
    for w in filtered:
        freq[w] = freq.get(w, 0) + 1

    sorted_words = sorted(freq.items(), key=lambda x: (-x[1], x[0]))
    return [w for w, _ in sorted_words[:top_n]]


def compute_ats_match(resume_text: str, job_description: str) -> dict:
    """Compare resume content against a job description for keyword overlap."""
    resume_lower = resume_text.lower()
    keywords = extract_keywords(job_description, top_n=35)

    matched = []
    missing = []
    for kw in keywords:
        if kw in resume_lower:
            matched.append(kw)
        else:
            missing.append(kw)

    total = len(keywords) if keywords else 1
    match_percent = round((len(matched) / total) * 100)

    resume_word_count = max(len(re.findall(r"\b\w+\b", resume_text)), 1)
    keyword_density = round((len(matched) / resume_word_count) * 100, 2)

    if match_percent >= 75:
        compatibility = "Excellent"
    elif match_percent >= 55:
        compatibility = "Good"
    elif match_percent >= 35:
        compatibility = "Fair"
    else:
        compatibility = "Poor"

    return {
        "matched": matched,
        "missing": missing,
        "match_percent": match_percent,
        "keyword_density": keyword_density,
        "compatibility": compatibility,
        "total_keywords": len(keywords),
    }


# ==============================================================================
# SECTION 6: RESUME REWRITE ENGINE (RULE-BASED BULLET IMPROVEMENT)
# ==============================================================================

WEAK_TO_STRONG = {
    "worked on": "engineered",
    "worked with": "collaborated with",
    "helped": "supported",
    "helped with": "contributed to",
    "was responsible for": "owned",
    "responsible for": "owned",
    "did": "executed",
    "made": "built",
    "handled": "managed",
    "in charge of": "led",
    "took care of": "managed",
    "involved in": "contributed to",
    "participated in": "collaborated on",
    "assisted with": "supported",
    "assisted in": "supported",
    "used": "leveraged",
    "dealt with": "resolved",
}

METRIC_SUGGESTIONS = [
    "Add a % improvement (e.g., increased efficiency)",
    "Add a dollar value or budget size",
    "Add a team size or number of stakeholders",
    "Add a timeframe (e.g., within 3 months)",
]