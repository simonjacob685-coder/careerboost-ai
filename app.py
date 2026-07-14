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

/* Hide default Streamlit chrome (safe, scoped selectors only) */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header[data-testid="stHeader"] {background: transparent;}

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
    "Add a % improvement (e.g., 'increased efficiency by 20%')",
    "Add a dollar value or budget size (e.g., 'managed a $50K budget')",
    "Add a team size or number of stakeholders involved",
    "Add a time saved metric (e.g., 'reduced processing time by 3 hours/week')",
    "Add the scale of users/customers impacted (e.g., 'served 10,000+ users')",
    "Add a ranking or comparison (e.g., 'ranked #1 among 15 teams')",
]


def is_bullet_candidate(line: str) -> bool:
    """
    Filter out lines that aren't real bullet/achievement statements —
    e.g. contact info, section headers, emails, phone numbers, single words,
    comma-separated skill lists, or degree/education lines.
    """
    if len(line.split()) < 3:
        return False
    lower = line.lower().strip()
    # Skip emails, phone numbers, URLs
    if re.search(r"@[\w\.-]+\.\w+", line):
        return False
    if re.search(r"\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}", line) and len(line.split()) <= 6:
        return False
    if re.search(r"https?://|www\.", lower):
        return False
    # Skip common section headers (short, all-caps-ish, no verbs)
    header_keywords = {"summary", "objective", "experience", "education", "skills",
                        "projects", "certifications", "contact", "profile", "employment history"}
    stripped_lower = lower.strip(":").strip()
    if stripped_lower in header_keywords:
        return False
    # Skip lines that are just a name / title (e.g., "John Doe") — no lowercase verbs, very short
    if len(line.split()) <= 4 and line.istitle():
        return False
    # Skip comma-separated skill/tool lists (e.g., "Python, JavaScript, SQL, AWS")
    comma_count = line.count(",")
    if comma_count >= 2 and len(line.split()) <= comma_count * 3:
        return False
    # Skip degree/education lines
    degree_keywords = {"bachelor", "master", "phd", "b.s.", "m.s.", "associate degree", "diploma"}
    if any(dk in lower for dk in degree_keywords):
        return False
    return True


def rewrite_bullet(bullet: str) -> dict:
    """
    Rewrite a single resume bullet point:
      - Replace weak phrases with stronger action verbs
      - Suggest a starting action verb if missing
      - Suggest a measurable achievement to add
    """
    original = bullet.strip()
    improved = original

    # Detect whether the bullet originally opened with a weak phrase
    # (so we know whether the replacement itself already fixed the opening verb)
    lower_original = original.lower()
    opened_with_weak_phrase = any(
        lower_original.startswith(weak) for weak in WEAK_TO_STRONG
    )

    # Replace weak phrases (case-insensitive) anywhere in the sentence
    for weak, strong in WEAK_TO_STRONG.items():
        pattern = re.compile(re.escape(weak), re.IGNORECASE)
        improved = pattern.sub(strong, improved)

    # Only prepend a new action verb if the sentence still doesn't start with one
    # AND the phrase-replacement step didn't already fix the opening word.
    first_word = improved.strip().split(" ")[0].lower().strip(".,") if improved.strip() else ""
    if first_word not in ACTION_VERBS and not opened_with_weak_phrase and improved:
        suggested_verb = random.choice(sorted(ACTION_VERBS))
        improved = f"{suggested_verb.capitalize()} {improved[0].lower()}{improved[1:]}" if len(improved) > 1 else improved

    # Capitalize first letter properly
    if improved:
        improved = improved[0].upper() + improved[1:]

    # Suggest a metric if none present
    has_metric = bool(re.search(r"\d", improved))
    metric_tip = None if has_metric else random.choice(METRIC_SUGGESTIONS)

    return {
        "original": original,
        "improved": improved,
        "metric_tip": metric_tip,
        "changed": improved.strip() != original.strip(),
    }


def rewrite_resume_bullets(text: str):
    """Split resume text into bullet-like lines and rewrite each."""
    raw_lines = clean_lines(text)
    bullets = [ln for ln in raw_lines if is_bullet_candidate(ln)]
    results = [rewrite_bullet(b) for b in bullets[:40]]  # cap for performance
    return results
# ==============================================================================
# SECTION 7: INTERVIEW COACH ENGINE
# ==============================================================================

HR_QUESTIONS = {
    "Easy": [
        "Tell me about yourself.",
        "Why do you want to work here?",
        "What are your strengths?",
        "What are your hobbies outside of work?",
        "How did you hear about this position?",
    ],
    "Medium": [
        "Why should we hire you over other candidates?",
        "What is your greatest weakness, and how are you improving it?",
        "Where do you see yourself in 5 years?",
        "How do you handle stress and pressure?",
        "Describe your ideal work environment.",
    ],
    "Hard": [
        "Tell me about a time you disagreed with your manager. How did you handle it?",
        "Why did you leave your last job / why are you leaving your current one?",
        "How do you handle receiving harsh criticism?",
        "What would your previous manager say is your biggest area for growth?",
        "Give an example of a time your work was criticized. What did you do?",
    ],
}

TECHNICAL_QUESTIONS = {
    "Easy": [
        "What tools/technologies are you most comfortable with?",
        "Explain a basic concept from your field in simple terms.",
        "Walk me through your typical workflow on a small task.",
        "What is the difference between two commonly confused terms in your field?",
        "How do you stay updated with trends in your industry?",
    ],
    "Medium": [
        "Describe a challenging technical/professional problem you solved recently.",
        "How would you approach a project with unclear requirements?",
        "Explain how you would optimize a process you've worked on before.",
        "What trade-offs do you consider when choosing between two approaches?",
        "How do you ensure quality and accuracy in your work?",
    ],
    "Hard": [
        "Design a system/process to solve [a complex problem in your field]. Walk me through your approach.",
        "How would you debug a critical issue with no clear documentation?",
        "Describe a time you had to learn a completely new skill under a tight deadline.",
        "How would you handle a situation where your technical recommendation was overruled?",
        "What is the most complex project you've led, and what made it complex?",
    ],
}

BEHAVIORAL_QUESTIONS = {
    "Easy": [
        "Tell me about a time you worked well in a team.",
        "Describe a goal you set and how you achieved it.",
        "Tell me about a time you helped a colleague.",
        "Describe a time you had to learn something quickly.",
        "Tell me about a project you're proud of.",
    ],
    "Medium": [
        "Tell me about a time you failed. What did you learn?",
        "Describe a situation where you had to manage multiple priorities.",
        "Tell me about a time you had to persuade someone to see your point of view.",
        "Describe a time you went above and beyond your job responsibilities.",
        "Tell me about a time you had to adapt to a major change.",
    ],
    "Hard": [
        "Describe a time you had a conflict with a team member and how you resolved it.",
        "Tell me about a time you made a mistake that impacted others. How did you handle it?",
        "Describe a time you had to make a difficult decision with incomplete information.",
        "Tell me about a time you had to lead a team through a crisis or setback.",
        "Describe a time you had to deliver bad news to a stakeholder or client.",
    ],
}

STAR_TIPS = [
    "**Situation** — Briefly set the scene: what was the context, when and where did this happen?",
    "**Task** — Explain your specific responsibility or the challenge you faced.",
    "**Action** — Describe the concrete steps YOU took (use 'I', not just 'we').",
    "**Result** — Share the outcome with measurable impact (numbers, %, time saved, etc.) and what you learned.",
]


def generate_interview_questions(difficulty: str, num_each: int = 3):
    """Return a dict of HR / Technical / Behavioral question lists for a given difficulty."""
    hr = random.sample(HR_QUESTIONS[difficulty], min(num_each, len(HR_QUESTIONS[difficulty])))
    tech = random.sample(TECHNICAL_QUESTIONS[difficulty], min(num_each, len(TECHNICAL_QUESTIONS[difficulty])))
    behavioral = random.sample(BEHAVIORAL_QUESTIONS[difficulty], min(num_each, len(BEHAVIORAL_QUESTIONS[difficulty])))
    return {"HR": hr, "Technical": tech, "Behavioral": behavioral}


# ==============================================================================
# SECTION 8: CAREER ROADMAP ENGINE (30-DAY PLAN)
# ==============================================================================

def generate_career_roadmap(scores: dict, ats_result: dict = None) -> dict:
    """Generate a personalized 30-day roadmap based on resume analysis results."""

    weak_spots = []
    if scores["weak_verbs"] > 3:
        weak_spots.append("replace weak verbs with strong action verbs")
    if scores["measurable_achievements"] < 3:
        weak_spots.append("add measurable, quantified achievements")
    if scores["passive_voice"] > 2:
        weak_spots.append("eliminate passive voice sentence structures")
    if scores["section_score"] < 80:
        weak_spots.append("add missing resume sections (Summary, Skills, Projects)")
    if not weak_spots:
        weak_spots.append("polish formatting and consistency across sections")

    missing_skills = ats_result["missing"][:6] if ats_result else [
        "industry-specific certifications", "portfolio projects", "leadership examples"
    ]

    week1 = {
        "title": "Week 1 — Resume Improvements",
        "items": [
            f"Focus area: {weak_spots[0]}.",
            "Rewrite your top 5 bullet points using the Resume Rewrite tool.",
            "Add 2-3 measurable achievements (numbers, %, $ impact).",
            "Ensure all required sections are present (Summary, Experience, Education, Skills).",
            "Proofread for grammar, tense consistency, and formatting.",
        ],
    }

    week2 = {
        "title": "Week 2 — Close Missing Skill Gaps",
        "items": [
            f"Target skills to develop: {', '.join(missing_skills[:4]) if missing_skills else 'core role-specific skills'}.",
            "Complete one free online course or tutorial in a missing skill area.",
            "Build or update one small project demonstrating that skill.",
            "Add the new skill/project to your resume and LinkedIn profile.",
            "Ask a mentor or peer for feedback on your updated resume.",
        ],
    }

    week3 = {
        "title": "Week 3 — Interview Preparation",
        "items": [
            "Practice 5 HR questions and 5 Behavioral questions using the STAR method.",
            "Practice 5 Technical/role-specific questions relevant to your target job.",
            "Record yourself answering 3 questions and review your delivery.",
            "Prepare 3 thoughtful questions to ask interviewers.",
            "Do one mock interview with a friend, mentor, or career coach.",
        ],
    }

    week4 = {
        "title": "Week 4 — Application Strategy",
        "items": [
            "Run your resume through the ATS Job Match tool for 3 target job postings.",
            "Customize your resume keywords for each application.",
            "Apply to 10-15 roles that align with your updated resume.",
            "Follow up on applications after 5-7 business days.",
            "Track applications, responses, and interview feedback in a spreadsheet.",
        ],
    }

    return {"week1": week1, "week2": week2, "week3": week3, "week4": week4}


# ==============================================================================
# SECTION 9: JOB SCAM DETECTOR ENGINE
# ==============================================================================

SCAM_RULES = [
    {
        "name": "Urgency Pressure Tactics",
        "patterns": [r"\burgent\b", r"immediate(ly)? start", r"limited (spots|slots|time)", r"act now", r"within 24 hours", r"hurry"],
        "explanation": "Legitimate employers rarely rush hiring decisions. Urgency is a classic pressure tactic to prevent you from researching the company.",
    },
    {
        "name": "Upfront Payment / Registration Fee Request",
        "patterns": [r"registration fee", r"processing fee", r"pay (a|an|the)? ?fee", r"training fee", r"deposit required", r"send money", r"purchase (a )?starter kit"],
        "explanation": "Real employers never ask candidates to pay money to get hired, for training materials, or for 'starter kits'. This is a major red flag for fraud.",
    },
    {
        "name": "WhatsApp / Telegram-Only Interview",
        "patterns": [r"whatsapp interview", r"contact (us|me) on whatsapp", r"telegram interview", r"interview via whatsapp", r"message us on telegram"],
        "explanation": "Scammers often move conversations to WhatsApp/Telegram to avoid traceable, official communication channels used by real companies.",
    },
    {
        "name": "Unrealistic Salary for Minimal Work",
        "patterns": [r"\$\d{3,}\s?/\s?(day|hour)", r"earn \$\d+.*no experience", r"easy money", r"get rich", r"guaranteed income", r"\$\d{4,}\s?(per week|weekly)"],
        "explanation": "Offers of very high pay for little to no experience or effort are a hallmark of scam or pyramid-scheme job postings.",
    },
    {
        "name": "Requests for Sensitive Personal/Financial Info Early",
        "patterns": [r"bank account details", r"social security number", r"ssn", r"routing number", r"credit card (number|info)", r"upload.*(passport|id) before interview"],
        "explanation": "Legitimate employers only request sensitive financial/personal information after a formal offer, not during initial contact.",
    },
    {
        "name": "Generic / Fake Recruiter Language",
        "patterns": [r"dear applicant", r"dear candidate", r"congratulations you.*selected", r"no interview (needed|required)", r"hired immediately", r"you have been selected"],
        "explanation": "Vague greetings and instant 'you're hired' messages without any interview process are common in mass-scam recruitment emails.",
    },
    {
        "name": "Suspicious / Free Email Domain for a 'Company'",
        "patterns": [r"@gmail\.com", r"@yahoo\.com", r"@hotmail\.com", r"@outlook\.com"],
        "explanation": "Official recruiters typically email from a company domain (e.g., @company.com), not free public email providers.",
    },
    {
        "name": "Phishing-Style Links or Wording",
        "patterns": [r"click here immediately", r"verify your account", r"confirm your details? here", r"update your (payment|banking) info", r"bit\.ly", r"tinyurl"],
        "explanation": "Shortened/suspicious links and urgent 'verify now' language are common phishing tactics used to steal credentials.",
    },
    {
        "name": "Vague Job Description",
        "patterns": [r"work from home.*flexible.*no experience", r"simple data entry.*high pay", r"be your own boss", r"unlimited earning potential"],
        "explanation": "Overly vague roles promising flexibility and high pay with no real job description often mask multi-level-marketing or scam schemes.",
    },
]


def detect_job_scam(text: str) -> dict:
    """Run all scam-detection rules against posting text and return findings."""
    lower_text = text.lower()
    findings = []

    for rule in SCAM_RULES:
        for pattern in rule["patterns"]:
            if re.search(pattern, lower_text):
                findings.append(rule)
                break  # only count each rule once

    risk_score = min(100, len(findings) * 18)

    if risk_score >= 70:
        risk_level = "HIGH RISK"
    elif risk_score >= 35:
        risk_level = "MODERATE RISK"
    elif risk_score > 0:
        risk_level = "LOW RISK"
    else:
        risk_level = "NO RED FLAGS DETECTED"

    return {"findings": findings, "risk_score": risk_score, "risk_level": risk_level}


# ==============================================================================
# SECTION 10: REPORT GENERATION (TXT + PDF)
# ==============================================================================

def build_txt_report(scores: dict, ats_result: dict = None) -> str:
    """Assemble a plain-text career report."""
    lines = []
    lines.append("=" * 60)
    lines.append("RESUMEBOOST AI PRO — CAREER ANALYSIS REPORT")
    lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("=" * 60)
    lines.append("")
    lines.append("OVERALL SCORES")
    lines.append("-" * 60)
    lines.append(f"Overall Score:        {scores['overall_score']} / 100")
    lines.append(f"Impact Score:         {scores['impact_score']} / 100")
    lines.append(f"ATS Score:            {scores['ats_score']} / 100")
    lines.append(f"Readability Score:    {scores['readability']} / 100")
    lines.append(f"Career Readiness:     {scores['career_readiness']} / 100")
    lines.append("")
    lines.append("RESUME STATISTICS")
    lines.append("-" * 60)
    lines.append(f"Word Count:               {scores['word_count']}")
    lines.append(f"Action Verbs Used:        {scores['action_verbs']}")
    lines.append(f"Weak Verbs Detected:      {scores['weak_verbs']}")
    lines.append(f"Filler Words Detected:    {scores['filler_words']}")
    lines.append(f"Passive Voice Instances:  {scores['passive_voice']}")
    lines.append(f"Measurable Achievements:  {scores['measurable_achievements']}")
    lines.append("")
    lines.append("RESUME SECTIONS DETECTED")
    lines.append("-" * 60)
    for section, present in scores["sections"].items():
        status = "PRESENT" if present else "MISSING"
        lines.append(f"{section:<20} {status}")
    lines.append("")

    if ats_result:
        lines.append("ATS JOB MATCH RESULTS")
        lines.append("-" * 60)
        lines.append(f"Match Percentage:   {ats_result['match_percent']}%")
        lines.append(f"Compatibility:      {ats_result['compatibility']}")
        lines.append(f"Keyword Density:    {ats_result['keyword_density']}%")
        lines.append(f"Matched Keywords:   {', '.join(ats_result['matched']) if ats_result['matched'] else 'None'}")
        lines.append(f"Missing Keywords:   {', '.join(ats_result['missing']) if ats_result['missing'] else 'None'}")
        lines.append("")

    lines.append("=" * 60)
    lines.append("Generated by ResumeBoost AI Pro — Next Byte Hacks V3")
    lines.append("100% Offline. 100% Rule-Based. Built with care.")
    lines.append("=" * 60)

    return "\n".join(lines)


def build_pdf_report(scores: dict, ats_result: dict = None) -> bytes:
    """Assemble a professional PDF career report using fpdf2."""
    if not FPDF_SUPPORT:
        return b""

    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # Header
    pdf.set_fill_color(79, 70, 229)
    pdf.rect(0, 0, 210, 30, "F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 20)
    pdf.set_xy(10, 8)
    pdf.cell(0, 10, "ResumeBoost AI Pro", ln=True)
    pdf.set_font("Helvetica", "", 11)
    pdf.set_xy(10, 18)
    pdf.cell(0, 8, "Career Analysis Report", ln=True)

    pdf.set_text_color(20, 20, 30)
    pdf.ln(20)

    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 10, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", ln=True)
    pdf.ln(2)

    def section_header(title):
        pdf.set_font("Helvetica", "B", 13)
        pdf.set_text_color(79, 70, 229)
        pdf.cell(0, 10, title, ln=True)
        pdf.set_text_color(20, 20, 30)
        pdf.set_font("Helvetica", "", 11)

    section_header("Overall Scores")
    score_rows = [
        ("Overall Score", scores["overall_score"]),
        ("Impact Score", scores["impact_score"]),
        ("ATS Score", scores["ats_score"]),
        ("Readability Score", scores["readability"]),
        ("Career Readiness", scores["career_readiness"]),
    ]
    for label, val in score_rows:
        pdf.cell(0, 8, f"{label}: {val} / 100", ln=True)
    pdf.ln(4)

    section_header("Resume Statistics")
    stat_rows = [
        ("Word Count", scores["word_count"]),
        ("Action Verbs Used", scores["action_verbs"]),
        ("Weak Verbs Detected", scores["weak_verbs"]),
        ("Filler Words Detected", scores["filler_words"]),
        ("Passive Voice Instances", scores["passive_voice"]),
        ("Measurable Achievements", scores["measurable_achievements"]),
    ]
    for label, val in stat_rows:
        pdf.cell(0, 8, f"{label}: {val}", ln=True)
    pdf.ln(4)

    section_header("Resume Sections Detected")
    for section, present in scores["sections"].items():
        status = "Present" if present else "Missing"
        pdf.cell(0, 8, f"{section}: {status}", ln=True)
    pdf.ln(4)

    if ats_result:
        section_header("ATS Job Match Results")
        pdf.cell(0, 8, f"Match Percentage: {ats_result['match_percent']}%", ln=True)
        pdf.cell(0, 8, f"Compatibility: {ats_result['compatibility']}", ln=True)
        pdf.cell(0, 8, f"Keyword Density: {ats_result['keyword_density']}%", ln=True)
        pdf.ln(2)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 8, "Matched Keywords:", ln=True)
        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(0, 6, ", ".join(ats_result["matched"]) if ats_result["matched"] else "None")
        pdf.ln(1)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 8, "Missing Keywords:", ln=True)
        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(0, 6, ", ".join(ats_result["missing"]) if ats_result["missing"] else "None")

    pdf.ln(8)
    pdf.set_font("Helvetica", "I", 9)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 8, "Generated by ResumeBoost AI Pro | Next Byte Hacks V3 | 100% Offline", ln=True)

    output = pdf.output(dest="S")
    if isinstance(output, str):
        return output.encode("latin-1", errors="ignore")
    return bytes(output)


# ==============================================================================
# SECTION 11: SMALL UI HELPER COMPONENTS
# ==============================================================================

def render_progress_bar(label: str, value: int, suffix: str = "/ 100"):
    """Render a custom animated gradient progress bar."""
    value = max(0, min(100, value))
    st.markdown(
        f"""
        <div style="margin-bottom:2px; font-weight:600; color:#E5E3F5; font-size:0.92rem;">
            {label} — <span style="color:#C4B5FD;">{value} {suffix}</span>
        </div>
        <div class="progress-outer">
            <div class="progress-inner" style="width:{value}%;"></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_score_card(label: str, value, col):
    """Render a single glass score card inside a given column."""
    with col:
        st.markdown(
            f"""
            <div class="score-card">
                <div class="score-value">{value}</div>
                <div class="score-label">{label}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_feature_card(icon: str, title: str, desc: str, col):
    with col:
        st.markdown(
            f"""
            <div class="gb-card">
                <div class="gb-feature-icon">{icon}</div>
                <div class="gb-feature-title">{title}</div>
                <div class="gb-feature-desc">{desc}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_section_title(title: str):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)


def render_footer():
    st.markdown(
        """
        <div class="gb-footer">
            🚀 <strong>ResumeBoost AI Pro</strong> — Built for <strong>Next Byte Hacks V3</strong><br>
            100% Offline · 100% Rule-Based Intelligence · No External APIs<br>
            Crafted with Python + Streamlit
        </div>
        """,
        unsafe_allow_html=True,
    )


def show_loading(message: str = "Analyzing your resume..."):
    with st.spinner(message):
        import time
        time.sleep(0.6)


# ==============================================================================
# SECTION 12: SESSION STATE INITIALIZATION
# ==============================================================================

if "resume_text" not in st.session_state:
    st.session_state.resume_text = ""
if "scores" not in st.session_state:
    st.session_state.scores = None
if "ats_result" not in st.session_state:
    st.session_state.ats_result = None
if "job_description" not in st.session_state:
    st.session_state.job_description = ""


# ==============================================================================
# SECTION 13: SIDEBAR NAVIGATION
# ==============================================================================

with st.sidebar:
    st.markdown(
        """
        <div style="text-align:center; padding: 0.6rem 0 1.2rem 0;">
            <div style="font-size:2.4rem;">🚀</div>
            <div style="font-family:'Poppins',sans-serif; font-weight:800; font-size:1.3rem; color:#F5F3FF;">
                ResumeBoost AI Pro
            </div>
            <div style="font-size:0.78rem; color:#A78BFA; font-weight:600;">
                Next Byte Hacks V3
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    page = st.radio(
        "Navigate",
        [
            "🏠 Home",
            "📄 Resume Analyzer",
            "🎯 ATS Job Match",
            "✍️ Resume Rewrite",
            "🎤 Interview Coach",
            "🗺️ Career Roadmap",
            "📊 Career Dashboard",
            "🛡️ Job Scam Detector",
            "📥 Reports",
        ],
        label_visibility="collapsed",
    )

    st.markdown("---")

    with st.expander("ℹ️ About"):
        st.markdown(
            "ResumeBoost AI Pro is an **offline AI career platform** that analyzes, "
            "improves, and prepares your resume and interview skills — no internet "
            "or external AI APIs required."
        )

    with st.expander("✨ Features"):
        st.markdown(
            "- Resume Analyzer\n"
            "- ATS Job Match\n"
            "- Resume Rewrite\n"
            "- Interview Coach\n"
            "- Career Roadmap\n"
            "- Career Dashboard\n"
            "- Job Scam Detector\n"
            "- TXT / PDF Reports"
        )

    with st.expander("🛠️ Technology"):
        st.markdown(
            "- **Python** + **Streamlit**\n"
            "- Rule-based NLP heuristics\n"
            "- `pdfplumber` for PDF parsing\n"
            "- `python-docx` for DOCX parsing\n"
            "- `fpdf2` for PDF report generation\n"
            "- Zero external APIs — fully offline"
        )

    with st.expander("👨‍💻 Developer"):
        st.markdown(
            "Built by a passionate developer team for **Next Byte Hacks V3**, "
            "focused on making career tools accessible without relying on paid "
            "AI APIs or internet connectivity."
        )

    with st.expander("🏆 Hackathon"):
        st.markdown(
            "**Event:** Next Byte Hacks V3\n\n"
            "**Category:** AI Career Tools\n\n"
            "**Highlight:** 100% offline intelligence — works anywhere, anytime."
        )

    st.markdown("---")
    st.caption("© 2026 ResumeBoost AI Pro. Built for hackathon demo purposes.")


# ==============================================================================
# SECTION 14: HOME PAGE
# ==============================================================================

def page_home():
    st.markdown(
        """
        <div class="hero-wrap">
            <div class="hero-badge">🏆 Next Byte Hacks V3 Submission</div>
            <div class="hero-title">🚀 ResumeBoost AI Pro</div>
            <div class="hero-sub">Land More Interviews. Build a Stronger Career.</div>
            <div style="color:rgba(255,255,255,0.85); font-size:0.98rem; max-width:640px;">
                A fully offline AI-powered career platform — analyze your resume, beat ATS filters,
                rewrite weak bullet points, practice interviews, and get a personalized 30-day
                roadmap to career success. No internet. No external APIs. 100% private.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_section_title("✨ Everything You Need to Get Hired")

    row1 = st.columns(3)
    render_feature_card("📄", "Resume Analysis", "Deep analysis of impact, tone, passive voice, filler words, and measurable achievements.", row1[0])
    render_feature_card("🎯", "ATS Checker", "Match your resume against any job description and uncover missing keywords instantly.", row1[1])
    render_feature_card("✍️", "Resume Rewrite", "Automatically strengthen weak bullet points with powerful action verbs.", row1[2])

    row2 = st.columns(3)
    render_feature_card("🎤", "Interview Coach", "Practice HR, technical, and behavioral questions with STAR method guidance.", row2[0])
    render_feature_card("🗺️", "Career Roadmap", "Get a personalized 30-day plan to improve your resume and land interviews.", row2[1])
    render_feature_card("🛡️", "Job Scam Detector", "Spot fraudulent job postings before you waste time — or worse, money.", row2[2])

    render_section_title("📊 Dashboard Preview")

    if st.session_state.scores:
        s = st.session_state.scores
        cols = st.columns(4)
        render_score_card("Overall", s["overall_score"], cols[0])
        render_score_card("ATS Ready", s["ats_score"], cols[1])
        render_score_card("Impact", s["impact_score"], cols[2])
        render_score_card("Readiness", s["career_readiness"], cols[3])
        st.caption("👆 Live scores from your uploaded resume. Visit the Career Dashboard tab for full detail.")
    else:
        cols = st.columns(4)
        render_score_card("Overall", "--", cols[0])
        render_score_card("ATS Ready", "--", cols[1])
        render_score_card("Impact", "--", cols[2])
        render_score_card("Readiness", "--", cols[3])
        st.info("💡 Upload your resume in the **Resume Analyzer** tab to unlock your live dashboard.")

    render_footer()


# ==============================================================================
# SECTION 15: RESUME ANALYZER PAGE
# ==============================================================================

def page_resume_analyzer():
    st.markdown(
        """
        <div class="hero-wrap" style="padding:2rem;">
            <div class="hero-title" style="font-size:2.1rem;">📄 Resume Analyzer</div>
            <div class="hero-sub" style="font-size:1rem;">Upload your resume for a full offline AI-style breakdown.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    input_mode = st.radio("Choose input method", ["Upload File", "Paste Text"], horizontal=True)

    resume_text = ""

    if input_mode == "Upload File":
        uploaded = st.file_uploader("Upload your resume (PDF, DOCX, or TXT)", type=["pdf", "docx", "txt"])
        if uploaded:
            with st.spinner("📖 Extracting text from your file..."):
                resume_text = extract_resume_text(uploaded)
            if not resume_text.strip():
                st.error("⚠️ We couldn't extract text from this file. Try a different file or paste text manually.")
            else:
                st.success(f"✅ Successfully extracted {len(resume_text.split())} words from **{uploaded.name}**")
    else:
        resume_text = st.text_area("Paste your resume text here", height=280, placeholder="Paste your full resume content here...")

    analyze_clicked = st.button("🔍 Analyze Resume", use_container_width=True)

    if analyze_clicked:
        if not resume_text or len(resume_text.strip()) < 30:
            st.error("⚠️ Please provide a valid resume with at least a few sentences before analyzing.")
        else:
            with st.spinner("🧠 Running rule-based AI analysis..."):
                import time
                time.sleep(0.8)
                scores = compute_resume_scores(resume_text)
            st.session_state.resume_text = resume_text
            st.session_state.scores = scores
            st.success("✅ Analysis complete! Scroll down to view your results.")

    if st.session_state.scores:
        s = st.session_state.scores
        render_section_title("📈 Score Overview")

        cols = st.columns(5)
        render_score_card("Overall", s["overall_score"], cols[0])
        render_score_card("Impact", s["impact_score"], cols[1])
        render_score_card("ATS", s["ats_score"], cols[2])
        render_score_card("Readability", s["readability"], cols[3])
        render_score_card("Readiness", s["career_readiness"], cols[4])

        st.write("")
        render_section_title("📊 Detailed Progress Bars")
        render_progress_bar("Overall Score", s["overall_score"])
        render_progress_bar("Impact Score", s["impact_score"])
        render_progress_bar("ATS Compatibility", s["ats_score"])
        render_progress_bar("Readability", s["readability"])
        render_progress_bar("Career Readiness", s["career_readiness"])

        render_section_title("🔎 Writing Quality Breakdown")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(
                f"""<div class="gb-card">
                <b>💪 Action Verbs</b><br>
                <span style="font-size:1.6rem; font-weight:800; color:#a7f3d0;">{s['action_verbs']}</span><br>
                <span style="color:#C4C1DD; font-size:0.85rem;">Strong verbs found in your resume</span>
                </div>""",
                unsafe_allow_html=True,
            )
        with c2:
            st.markdown(
                f"""<div class="gb-card">
                <b>⚠️ Weak Verbs</b><br>
                <span style="font-size:1.6rem; font-weight:800; color:#fca5a5;">{s['weak_verbs']}</span><br>
                <span style="color:#C4C1DD; font-size:0.85rem;">Phrases that dilute your impact</span>
                </div>""",
                unsafe_allow_html=True,
            )
        with c3:
            st.markdown(
                f"""<div class="gb-card">
                <b>📉 Passive Voice</b><br>
                <span style="font-size:1.6rem; font-weight:800; color:#fbbf24;">{s['passive_voice']}</span><br>
                <span style="color:#C4C1DD; font-size:0.85rem;">Instances of passive sentence structure</span>
                </div>""",
                unsafe_allow_html=True,
            )

        c4, c5, c6 = st.columns(3)
        with c4:
            st.markdown(
                f"""<div class="gb-card">
                <b>🧹 Filler Words</b><br>
                <span style="font-size:1.6rem; font-weight:800; color:#fca5a5;">{s['filler_words']}</span><br>
                <span style="color:#C4C1DD; font-size:0.85rem;">Vague or unnecessary words</span>
                </div>""",
                unsafe_allow_html=True,
            )
        with c5:
            st.markdown(
                f"""<div class="gb-card">
                <b>📐 Measurable Wins</b><br>
                <span style="font-size:1.6rem; font-weight:800; color:#a7f3d0;">{s['measurable_achievements']}</span><br>
                <span style="color:#C4C1DD; font-size:0.85rem;">Quantified achievements detected</span>
                </div>""",
                unsafe_allow_html=True,
            )
        with c6:
            st.markdown(
                f"""<div class="gb-card">
                <b>📝 Word Count</b><br>
                <span style="font-size:1.6rem; font-weight:800; color:#c4b5fd;">{s['word_count']}</span><br>
                <span style="color:#C4C1DD; font-size:0.85rem;">Ideal range: 400 – 800 words</span>
                </div>""",
                unsafe_allow_html=True,
            )

        render_section_title("🧩 Resume Sections Detected")
        section_cols = st.columns(3)
        for i, (section, present) in enumerate(s["sections"].items()):
            with section_cols[i % 3]:
                badge = '<span class="badge-good">✔ Present</span>' if present else '<span class="badge-bad">✘ Missing</span>'
                st.markdown(
                    f"""<div class="gb-card" style="text-align:center;">
                    <b>{section}</b><br><br>{badge}
                    </div>""",
                    unsafe_allow_html=True,
                )

        render_section_title("💡 Recommendations")
        recs = []
        if s["weak_verbs"] > 2:
            recs.append("Replace weak verbs like *'helped'* or *'responsible for'* with strong action verbs like *'led'* or *'delivered'*.")
        if s["passive_voice"] > 1:
            recs.append("Rewrite passive sentences (e.g., *'was tasked with'*) into active voice (e.g., *'led'*, *'managed'*).")
        if s["measurable_achievements"] < 3:
            recs.append("Add more numbers — percentages, dollar values, or team sizes — to quantify your impact.")
        if s["filler_words"] > 2:
            recs.append("Remove filler phrases like *'very'*, *'hard worker'*, or *'team player'* — show it, don't say it.")
        if s["section_score"] < 100:
            missing_sections = [k for k, v in s["sections"].items() if not v]
            recs.append(f"Add these missing sections: **{', '.join(missing_sections)}**.")
        if s["word_count"] < 200:
            recs.append("Your resume looks short — consider expanding your experience and project details.")
        if not recs:
            recs.append("🎉 Great job! Your resume is well-optimized. Fine-tune wording for even more polish.")

        for r in recs:
            st.markdown(f"- {r}")

    render_footer()


# ==============================================================================
# SECTION 16: ATS JOB MATCH PAGE
# ==============================================================================

def page_ats_match():
    st.markdown(
        """
        <div class="hero-wrap" style="padding:2rem;">
            <div class="hero-title" style="font-size:2.1rem;">🎯 ATS Job Match</div>
            <div class="hero-sub" style="font-size:1rem;">Paste a job description to see how well your resume matches.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not st.session_state.resume_text:
        st.warning("⚠️ Please analyze a resume first in the **Resume Analyzer** tab.")

    job_description = st.text_area(
        "Paste the job description here",
        height=220,
        value=st.session_state.job_description,
        placeholder="Paste the full job posting text here...",
    )

    match_clicked = st.button("🎯 Run ATS Match", use_container_width=True)

    if match_clicked:
        if not st.session_state.resume_text:
            st.error("⚠️ No resume found. Please analyze your resume first.")
        elif not job_description or len(job_description.strip()) < 30:
            st.error("⚠️ Please paste a valid job description (at least a few sentences).")
        else:
            with st.spinner("🧠 Matching keywords against your resume..."):
                import time
                time.sleep(0.7)
                result = compute_ats_match(st.session_state.resume_text, job_description)
            st.session_state.ats_result = result
            st.session_state.job_description = job_description
            st.success("✅ ATS match analysis complete!")

    if st.session_state.ats_result:
        r = st.session_state.ats_result
        render_section_title("📊 Match Overview")

        cols = st.columns(4)
        render_score_card("Match %", f"{r['match_percent']}%", cols[0])
        render_score_card("Compatibility", r["compatibility"], cols[1])
        render_score_card("Keyword Density", f"{r['keyword_density']}%", cols[2])
        render_score_card("Total Keywords", r["total_keywords"], cols[3])

        render_progress_bar("Overall Keyword Match", r["match_percent"], suffix="%")

        render_section_title("✅ Matched Keywords")
        if r["matched"]:
            chips = "".join([f'<span class="chip-match">✔ {kw}</span>' for kw in r["matched"]])
            st.markdown(chips, unsafe_allow_html=True)
        else:
            st.info("No matched keywords found. Consider revising your resume to align with this role.")

        render_section_title("❌ Missing Keywords")
        if r["missing"]:
            chips = "".join([f'<span class="chip-missing">✘ {kw}</span>' for kw in r["missing"]])
            st.markdown(chips, unsafe_allow_html=True)
            st.warning("Adding these keywords (where truthful and relevant) can significantly improve ATS pass-through rates.")
        else:
            st.success("🎉 No missing keywords — excellent alignment with this job description!")

        render_section_title("📘 Why Keywords Matter")
        st.markdown(
            """
            <div class="gb-card">
            Applicant Tracking Systems (ATS) scan resumes for keywords found in the job description
            before a human ever sees them. Resumes with low keyword overlap are often auto-filtered out —
            regardless of your real qualifications. Matching relevant keywords (skills, tools, certifications,
            and role-specific terms) increases your chances of reaching a recruiter's inbox.
            <br><br>
            <b>Tip:</b> Only add keywords that are truthful and reflect your real experience. Keyword-stuffing
            without substance can hurt you in the actual interview.
            </div>
            """,
            unsafe_allow_html=True,
        )

    render_footer()


# ==============================================================================
# SECTION 17: RESUME REWRITE PAGE
# ==============================================================================

def page_resume_rewrite():
    st.markdown(
        """
        <div class="hero-wrap" style="padding:2rem;">
            <div class="hero-title" style="font-size:2.1rem;">✍️ Resume Rewrite</div>
            <div class="hero-sub" style="font-size:1rem;">Turn weak bullet points into powerful, achievement-driven statements.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    source_text = st.session_state.resume_text
    st.info("💡 This tool rewrites weak bullet points using rule-based logic — strong verbs, better phrasing, and metric suggestions.")

    manual_text = st.text_area(
        "Resume content to rewrite (auto-filled from Resume Analyzer if available)",
        value=source_text,
        height=220,
        placeholder="Paste resume bullet points here, one per line...",
    )

    rewrite_clicked = st.button("✨ Rewrite My Bullet Points", use_container_width=True)

    if rewrite_clicked:
        if not manual_text or len(manual_text.strip()) < 10:
            st.error("⚠️ Please provide some resume text or bullet points to rewrite.")
        else:
            with st.spinner("✍️ Rewriting bullet points with rule-based AI logic..."):
                import time
                time.sleep(0.8)
                results = rewrite_resume_bullets(manual_text)

            if not results:
                st.warning("No bullet-style lines detected. Try pasting individual resume lines.")
            else:
                render_section_title(f"🪄 Rewritten Bullet Points ({len(results)} found)")
                changed_count = sum(1 for r in results if r["changed"])
                st.success(f"✅ Improved {changed_count} out of {len(results)} bullet points.")

                for i, res in enumerate(results, 1):
                    st.markdown(f"**Bullet {i}**")
                    st.markdown(f'<div class="rewrite-before">🔴 Before: {res["original"]}</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="rewrite-after">🟢 After: {res["improved"]}</div>', unsafe_allow_html=True)
                    if res["metric_tip"]:
                        st.caption(f"💡 Suggestion: {res['metric_tip']}")
                    st.write("")

    render_footer()


# ==============================================================================
# SECTION 18: INTERVIEW COACH PAGE
# ==============================================================================

def page_interview_coach():
    st.markdown(
        """
        <div class="hero-wrap" style="padding:2rem;">
            <div class="hero-title" style="font-size:2.1rem;">🎤 Interview Coach</div>
            <div class="hero-sub" style="font-size:1rem;">Practice HR, technical, and behavioral interview questions.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    difficulty = st.select_slider("Select difficulty level", options=["Easy", "Medium", "Hard"], value="Medium")

    generate_clicked = st.button("🎲 Generate Interview Questions", use_container_width=True)

    if generate_clicked:
        with st.spinner("🎤 Preparing your personalized interview set..."):
            import time
            time.sleep(0.6)
            questions = generate_interview_questions(difficulty, num_each=4)
        st.session_state.interview_questions = questions
        st.session_state.interview_difficulty = difficulty

    if "interview_questions" in st.session_state:
        q = st.session_state.interview_questions
        diff = st.session_state.get("interview_difficulty", difficulty)

        render_section_title(f"🗂️ {diff} Difficulty — Question Set")

        tab1, tab2, tab3, tab4 = st.tabs(["👥 HR Questions", "🧠 Technical Questions", "🎭 Behavioral Questions", "⭐ STAR Method Tips"])

        with tab1:
            for i, question in enumerate(q["HR"], 1):
                st.markdown(f'<div class="gb-card"><b>Q{i}.</b> {question}</div>', unsafe_allow_html=True)

        with tab2:
            for i, question in enumerate(q["Technical"], 1):
                st.markdown(f'<div class="gb-card"><b>Q{i}.</b> {question}</div>', unsafe_allow_html=True)

        with tab3:
            for i, question in enumerate(q["Behavioral"], 1):
                st.markdown(f'<div class="gb-card"><b>Q{i}.</b> {question}</div>', unsafe_allow_html=True)

        with tab4:
            st.markdown("Use the **STAR method** to structure clear, compelling answers to behavioral questions:")
            for tip in STAR_TIPS:
                st.markdown(f'<div class="gb-card">{tip}</div>', unsafe_allow_html=True)
            st.info("💡 Tip: Always end your STAR answer with a measurable result or key learning.")
    else:
        st.info("👆 Select a difficulty level and click **Generate Interview Questions** to begin practicing.")

    render_footer()


# ==============================================================================
# SECTION 19: CAREER ROADMAP PAGE
# ==============================================================================

def page_career_roadmap():
    st.markdown(
        """
        <div class="hero-wrap" style="padding:2rem;">
            <div class="hero-title" style="font-size:2.1rem;">🗺️ Career Roadmap</div>
            <div class="hero-sub" style="font-size:1rem;">Your personalized 30-day plan to become interview-ready.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not st.session_state.scores:
        st.warning("⚠️ Please analyze your resume first in the **Resume Analyzer** tab for a personalized roadmap.")
        st.info("Showing a general roadmap template below.")
        dummy_scores = compute_resume_scores(
            "Responsible for helping the team with various tasks. Worked on projects."
        )
        roadmap = generate_career_roadmap(dummy_scores, st.session_state.ats_result)
    else:
        roadmap = generate_career_roadmap(st.session_state.scores, st.session_state.ats_result)

    render_section_title("📅 Your 30-Day Career Success Plan")

    for week_key in ["week1", "week2", "week3", "week4"]:
        week = roadmap[week_key]
        items_html = "".join([f"<li style='margin-bottom:6px;'>{item}</li>" for item in week["items"]])
        st.markdown(
            f"""
            <div class="week-card">
                <div class="week-title">{week['title']}</div>
                <ul style="color:#DDD9F5; font-size:0.92rem; padding-left:1.2rem;">
                    {items_html}
                </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.success("🎯 Consistency beats intensity — small daily progress compounds into big career wins!")
    render_footer()


# ==============================================================================
# SECTION 20: CAREER DASHBOARD PAGE
# ==============================================================================

def page_career_dashboard():
    st.markdown(
        """
        <div class="hero-wrap" style="padding:2rem;">
            <div class="hero-title" style="font-size:2.1rem;">📊 Career Dashboard</div>
            <div class="hero-sub" style="font-size:1rem;">A complete snapshot of your career readiness.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not st.session_state.scores:
        st.warning("⚠️ Please analyze your resume first in the **Resume Analyzer** tab to populate your dashboard.")
        render_footer()
        return

    s = st.session_state.scores

    # Derived composite metrics for a richer dashboard
    communication_score = max(0, min(100, round(100 - (s["filler_words"] * 5) - (s["passive_voice"] * 4))))
    confidence_score = max(0, min(100, round((s["action_verbs"] * 4) + (s["measurable_achievements"] * 3))))

    render_section_title("🏆 Career Readiness Overview")

    cols = st.columns(3)
    render_score_card("Resume Quality", s["overall_score"], cols[0])
    render_score_card("ATS Readiness", s["ats_score"], cols[1])
    render_score_card("Communication", communication_score, cols[2])

    cols2 = st.columns(3)
    render_score_card("Confidence Signal", confidence_score, cols2[0])
    render_score_card("Career Readiness", s["career_readiness"], cols2[1])
    render_score_card("Overall Score", s["overall_score"], cols2[2])

    render_section_title("📈 Visual Breakdown")
    render_progress_bar("Resume Quality", s["overall_score"])
    render_progress_bar("ATS Readiness", s["ats_score"])
    render_progress_bar("Communication", communication_score)
    render_progress_bar("Confidence Signal", confidence_score)
    render_progress_bar("Career Readiness", s["career_readiness"])

    render_section_title("📉 Score Comparison Chart")
    comparison_data = [
        ("Resume Quality", s["overall_score"]),
        ("ATS Readiness", s["ats_score"]),
        ("Communication", communication_score),
        ("Confidence", confidence_score),
        ("Career Readiness", s["career_readiness"]),
    ]
    for metric_name, metric_value in comparison_data:
        render_progress_bar(metric_name, metric_value)

    if s["overall_score"] >= 80:
        st.success("🌟 You're in excellent shape! Fine-tune small details and start applying with confidence.")
    elif s["overall_score"] >= 55:
        st.warning("📈 Solid foundation — a few targeted improvements will meaningfully boost your results.")
    else:
        st.error("🔧 Your resume needs focused work. Check out the Resume Rewrite and Career Roadmap tools.")

    render_footer()


# ==============================================================================
# SECTION 21: JOB SCAM DETECTOR PAGE
# ==============================================================================

def page_scam_detector():
    st.markdown(
        """
        <div class="hero-wrap" style="padding:2rem;">
            <div class="hero-title" style="font-size:2.1rem;">🛡️ Job Scam Detector</div>
            <div class="hero-sub" style="font-size:1rem;">Paste a job posting or recruiter message to check for red flags.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    posting_text = st.text_area(
        "Paste the job posting or recruiter message here",
        height=220,
        placeholder="Paste the suspicious job posting, email, or message text here...",
    )

    scan_clicked = st.button("🔍 Scan for Red Flags", use_container_width=True)

    if scan_clicked:
        if not posting_text or len(posting_text.strip()) < 15:
            st.error("⚠️ Please paste some job posting text to scan.")
        else:
            with st.spinner("🛡️ Scanning for scam indicators..."):
                import time
                time.sleep(0.7)
                result = detect_job_scam(posting_text)
            st.session_state.scam_result = result

    if "scam_result" in st.session_state:
        r = st.session_state.scam_result

        render_section_title("🚨 Risk Assessment")

        risk_colors = {
            "HIGH RISK": "#ef4444",
            "MODERATE RISK": "#f59e0b",
            "LOW RISK": "#fbbf24",
            "NO RED FLAGS DETECTED": "#10b981",
        }
        color = risk_colors.get(r["risk_level"], "#A78BFA")

        st.markdown(
            f"""
            <div class="gb-card" style="text-align:center; border: 2px solid {color};">
                <div style="font-size:1.1rem; color:#C4C1DD; font-weight:600;">Risk Level</div>
                <div style="font-size:2rem; font-weight:800; color:{color};">{r['risk_level']}</div>
                <div style="color:#C4C1DD; margin-top:6px;">Risk Score: {r['risk_score']} / 100</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        render_progress_bar("Scam Risk Score", r["risk_score"])

        render_section_title("🔎 Detected Warning Signs")
        if r["findings"]:
            for finding in r["findings"]:
                st.markdown(
                    f"""
                    <div class="alert-danger">
                        <b>⚠️ {finding['name']}</b><br>
                        {finding['explanation']}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
        else:
            st.markdown(
                """
                <div class="alert-safe">
                    <b>✅ No major red flags detected.</b><br>
                    This posting doesn't match common scam patterns, but always verify company
                    legitimacy independently (official website, LinkedIn presence, employee reviews)
                    before sharing personal information.
                </div>
                """,
                unsafe_allow_html=True,
            )

        render_section_title("🧭 General Safety Tips")
        st.markdown(
            """
            <div class="gb-card">
            - Never pay money to apply for, interview for, or "secure" a job.<br>
            - Verify recruiter identities via LinkedIn and official company domains.<br>
            - Be cautious of jobs promising high pay for minimal effort or experience.<br>
            - Avoid sharing sensitive financial/personal information before a formal offer.<br>
            - Trust your instincts — if something feels off, research before proceeding.
            </div>
            """,
            unsafe_allow_html=True,
        )

    render_footer()


# ==============================================================================
# SECTION 22: REPORTS PAGE
# ==============================================================================

def page_reports():
    st.markdown(
        """
        <div class="hero-wrap" style="padding:2rem;">
            <div class="hero-title" style="font-size:2.1rem;">📥 Reports</div>
            <div class="hero-sub" style="font-size:1rem;">Download your full career analysis as TXT or PDF.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not st.session_state.scores:
        st.warning("⚠️ Please analyze your resume first in the **Resume Analyzer** tab to generate a report.")
        render_footer()
        return

    s = st.session_state.scores
    ats = st.session_state.ats_result

    render_section_title("📋 Report Preview")
    st.markdown(
        f"""
        <div class="gb-card">
        <b>Overall Score:</b> {s['overall_score']} / 100 {score_badge(s['overall_score'])}<br>
        <b>ATS Score:</b> {s['ats_score']} / 100 {score_badge(s['ats_score'])}<br>
        <b>Impact Score:</b> {s['impact_score']} / 100 {score_badge(s['impact_score'])}<br>
        <b>Readability:</b> {s['readability']} / 100 {score_badge(s['readability'])}<br>
        <b>Career Readiness:</b> {s['career_readiness']} / 100 {score_badge(s['career_readiness'])}<br>
        {"<b>ATS Match:</b> " + str(ats['match_percent']) + "% (" + ats['compatibility'] + ")" if ats else ""}
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_section_title("⬇️ Download Your Reports")

    col1, col2 = st.columns(2) 

    with col1:
        txt_report = build_txt_report(s, ats)
        st.download_button(
            label="📄 Download TXT Report",
            data=txt_report,
            file_name=f"ResumeBoost_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
            mime="text/plain",
            use_container_width=True,
        )

    with col2:
        if FPDF_SUPPORT:
            with st.spinner("🖨️ Building professional PDF report..."):
                pdf_bytes = build_pdf_report(s, ats)
            if pdf_bytes:
                st.download_button(
                    label="📕 Download PDF Report",
                    data=pdf_bytes,
                    file_name=f"ResumeBoost_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            else:
                st.error("⚠️ PDF generation failed. Please try the TXT report instead.")
        else:
            st.info("📕 PDF export requires the `fpdf2` package. Install it via requirements.txt.")

    render_footer()


# ==============================================================================
# SECTION 23: PAGE ROUTER
# ==============================================================================

PAGE_ROUTER = {
    "🏠 Home": page_home,
    "📄 Resume Analyzer": page_resume_analyzer,
    "🎯 ATS Job Match": page_ats_match,
    "✍️ Resume Rewrite": page_resume_rewrite,
    "🎤 Interview Coach": page_interview_coach,
    "🗺️ Career Roadmap": page_career_roadmap,
    "📊 Career Dashboard": page_career_dashboard,
    "🛡️ Job Scam Detector": page_scam_detector,
    "📥 Reports": page_reports,
}

# Route to the selected page with basic error handling for robustness
try:
    PAGE_ROUTER[page]()
except Exception as e:
    st.error(f"⚠️ Something went wrong while loading this page: {e}")
    st.info("Please try refreshing the app or navigating to a different section.")
