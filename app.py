"""
LexPilot India — Contract Intelligence Platform (v1.0, single file)

Run:
    pip install -r requirements.txt
    streamlit run app.py

First launch asks you to create the administrator account.
Optional: GEMINI_API_KEY (secrets.toml or environment variable) for AI drafting.
Optional: Tesseract OCR installed on the machine for scanned PDFs.
Optional: LEXPILOT_DB environment variable to store the database somewhere durable.
"""

import difflib
import hashlib
import hmac
import io
import json
import os
import re
import secrets
import sqlite3
import tempfile
import time
from contextlib import contextmanager
from datetime import datetime, timedelta
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, RGBColor

# ================================================================ CONSTANTS
APP_VERSION = "1.0"
APP_DIR = Path(__file__).parent
DB_PATH = Path(os.getenv("LEXPILOT_DB", str(APP_DIR / "lexpilot.db")))

MAX_UPLOAD_MB = 25
MAX_PAGES = 500
MAX_OCR_PAGES = 60
MAX_CHARS = 400_000          # contract text sent to the AI (the UI warns if text is cut)
SESSION_TIMEOUT_MIN = 60
MAX_FAILED_LOGINS = 5
LOCK_MINUTES = 10
AI_HOURLY_LIMIT = 40
PBKDF2_ITER = 240_000
MIN_PW_LEN = 10

ROLES = ["staff", "advocate", "admin"]
ROLE_LABEL = {"staff": "Staff", "advocate": "Advocate", "admin": "Administrator"}
PERMS = {
    "review": {"advocate", "admin"},
    "playbook_edit": {"advocate", "admin"},
    "delete_matter": {"advocate", "admin"},
}

RISK_LEVELS = ["High", "Review", "Informational", "Unassessed"]
APPROVED = "Reviewed — approved for delivery"
REVIEW_OPTIONS = ["Pending", "Review in progress", APPROVED]
DRAFT_UNREVIEWED = "AI draft — not reviewed"
DRAFT_REVIEWED = "Reviewed by advocate"
CONTRACT_TYPES = [
    "Vendor agreement", "Services agreement", "NDA", "SaaS / technology agreement",
    "Consulting agreement", "Employment agreement", "Distribution agreement", "Lease / licence", "Other",
]
PAGES_BASE = ["Overview", "New contract review", "Matter workspace", "Playbooks",
              "Compare versions", "My account", "Product roadmap"]

st.set_page_config(
    page_title="LexPilot | Contract Intelligence",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================ BRAND / DESIGN
LOGO_SVG = """<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="LexPilot mark">
<defs><linearGradient id="lpg" x1="0" y1="0" x2="1" y2="1">
<stop offset="0" stop-color="#f1d39e"/><stop offset="1" stop-color="#c28d42"/></linearGradient></defs>
<rect x="3" y="3" width="94" height="94" rx="25" fill="#f7f0e4" fill-opacity=".10" stroke="#e4c38e" stroke-opacity=".55"/>
<path d="M50 17v54M32 31h36M23 42l-10 20h20L23 42Zm54 0L67 62h20L77 42ZM35 73h30" fill="none" stroke="url(#lpg)" stroke-width="4.2" stroke-linecap="round" stroke-linejoin="round"/>
<path d="M50 17l-5 7h10l-5-7Z" fill="#f1d39e"/><circle cx="50" cy="82" r="4" fill="#f1d39e"/></svg>"""

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root { --ink:#17243b; --muted:#718096; --navy:#101d33; --gold:#d9ad63; --line:#e7ebf1; }
html, body, [class*="css"] { font-family:'DM Sans',sans-serif; }
.stApp { background: radial-gradient(circle at 88% 2%, rgba(217,173,99,.09), transparent 23%), #f6f8fb; }
.block-container { max-width:1480px; padding-top:1.25rem; padding-bottom:3rem; }
[data-testid="stSidebar"] { background:linear-gradient(180deg,#101d33 0%,#172944 100%); }
[data-testid="stSidebar"] * { color:#edf2f8 !important; }
[data-testid="stSidebar"] hr { border-color:rgba(255,255,255,.15); }
[data-testid="stSidebar"] .stButton>button { background:rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.28); }
h1,h2,h3,h4 { font-family:'Manrope',sans-serif !important; color:var(--ink); letter-spacing:-.025em; }
h1 { font-weight:800 !important; }
.hero { position:relative; overflow:hidden; border-radius:22px; padding:27px 30px;
  background:linear-gradient(115deg,#101d33 0%,#1d3558 65%,#28456c 100%);
  color:#fff; box-shadow:0 16px 42px rgba(18,34,58,.15); margin-bottom:22px; }
.hero:before { content:""; position:absolute; width:330px; height:330px; right:-80px; top:-190px;
  border:1px solid rgba(232,194,128,.35); border-radius:50%;
  box-shadow:0 0 0 28px rgba(232,194,128,.045),0 0 0 58px rgba(232,194,128,.035);
  animation:orbit 18s linear infinite; }
.hero:after { content:""; position:absolute; width:180px; height:180px; right:18%; bottom:-145px;
  border-radius:50%; background:rgba(217,173,99,.16); animation:pulse 5s ease-in-out infinite; }
@keyframes orbit { from{transform:rotate(0deg)} to{transform:rotate(360deg)} }
@keyframes pulse { 0%,100%{transform:scale(.9);opacity:.45} 50%{transform:scale(1.15);opacity:.9} }
@media (prefers-reduced-motion: reduce) { .hero:before,.hero:after { animation:none; } }
.hero-inner { position:relative; z-index:1; display:flex; align-items:center; gap:17px; }
.hero h1 { color:#fff !important; font-size:2.05rem; margin:0; }
.hero p { color:#d9e2ef; margin:.35rem 0 0 0; font-size:.98rem; }
.eyebrow { text-transform:uppercase; letter-spacing:.16em; color:#e4c38e; font-size:.68rem; font-weight:700; }
.logo-mark { width:54px; height:54px; flex:0 0 54px; }
.metric-card { border:1px solid #e8edf3; border-radius:16px; background:rgba(255,255,255,.92);
  padding:17px 18px; box-shadow:0 5px 20px rgba(24,43,70,.035); min-height:112px; }
.metric-label { color:#7a8799; font-size:.78rem; font-weight:600; }
.metric-value { color:#17243b; font-family:'Manrope'; font-size:1.75rem; font-weight:800; margin-top:5px; }
.metric-foot { color:#8793a3; font-size:.72rem; margin-top:1px; }
.section-kicker { color:#a17a3e; font-size:.72rem; letter-spacing:.13em; text-transform:uppercase; font-weight:800; }
.soft-card { border:1px solid #e8edf3; background:#fff; border-radius:17px; padding:18px 20px;
  box-shadow:0 5px 18px rgba(20,40,70,.035); }
.small-muted { color:#8793a3; font-size:.82rem; }
.pill { display:inline-block; border-radius:30px; padding:4px 9px; font-size:.72rem; font-weight:700; background:#eef2f7; color:#536174; }
.pill-gold { background:#fbf2e3; color:#8c642d; }
.pill-green { background:#e8f6ef; color:#24754b; }
.pill-red { background:#fff0ef; color:#a63e38; }
.stButton>button, .stDownloadButton>button { border-radius:10px; font-weight:700; border:1px solid #dce3ec;
  min-height:2.65rem; transition:transform .15s ease, box-shadow .15s ease; }
.stButton>button:hover,.stDownloadButton>button:hover { transform:translateY(-1px); box-shadow:0 7px 18px rgba(25,43,69,.10); }
.stButton>button[kind="primary"], .stFormSubmitButton>button[kind="primary"] {
  background:linear-gradient(110deg,#172a47,#29486f); border:0; color:#fff; }
div[data-testid="stExpander"] { border:1px solid #e7ebf1; border-radius:13px; background:#fff; }
div[data-testid="stTabs"] button { font-weight:700; }
.footer-note { color:#8a96a6; font-size:.75rem; text-align:center; padding:22px 0 0; }
@media(max-width:700px) { .hero { padding:22px 20px; } .hero h1 { font-size:1.6rem; } }
</style>
"""


def render_header():
    st.markdown(
        f"""
    <div class="hero"><div class="hero-inner">
      <div class="logo-mark">{LOGO_SVG}</div>
      <div>
        <div class="eyebrow">Legal operations · India</div>
        <h1>LexPilot</h1>
        <p>Contract intelligence for growing businesses — with a human review workflow.</p>
      </div></div></div>""",
        unsafe_allow_html=True,
    )


st.markdown(CSS, unsafe_allow_html=True)


# ================================================================== DATABASE
@contextmanager
def db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def now():
    return datetime.now().isoformat(timespec="seconds")


def ensure_columns(conn, table, cols):
    have = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}
    for col, decl in cols.items():
        if col not in have:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {decl}")


PLAYBOOK_DEFAULTS = [
    # (applies to, clause, detect patterns, red-flag patterns, required, risk if missing, preferred, fallback)
    ("Vendor agreement;Services agreement;SaaS / technology agreement;Consulting agreement;Distribution agreement",
     "Limitation of liability",
     [r"limitation of liability", r"aggregate liability", r"liabilit.{0,60}(?:shall not exceed|limited to|capped)",
      r"not be liable for (?:any )?(?:indirect|consequential)"],
     [r"unlimited liability", r"no limit.{0,30}liabilit", r"liabilit.{0,30}unlimited"],
     1, "High",
     "Mutual liability cap linked to fees paid or payable; exclusion of indirect and consequential loss; carve-outs kept narrow and expressly listed.",
     "Higher or separate cap for confidentiality and data breaches; carve-outs limited to fraud and wilful misconduct."),
    ("All", "Indemnity",
     [r"indemnif", r"hold harmless"],
     [r"indemnif[^.]{0,150}(?:any and all|all losses|whatsoever)"],
     0, "Review",
     "Mutual; limited to third-party claims; notice, defence-control and mitigation provisions; subject to the liability cap or an agreed super-cap.",
     "One-way indemnity only where the risk allocation is commercially justified, with a defined cap."),
    ("All", "Termination",
     [r"terminat"],
     [r"terminat[^.]{0,100}(?:sole discretion|without (?:any )?(?:cause|notice|reason))"],
     1, "Review",
     "Mutual termination for material breach with a cure period; reasonable notice for convenience termination; payment for work done; transition support; survival clause.",
     "Asymmetric convenience termination only with longer notice and compensation for committed costs."),
    ("All", "Governing law and dispute resolution",
     [r"governing law", r"jurisdiction", r"arbitrat", r"courts? (?:at|of|in)"],
     [r"laws? of (?:england|new york|singapore|delaware|california|the state of)"],
     1, "Review",
     "Laws of India; clear seat or venue; agreed dispute procedure (escalation, then arbitration or courts); consistent clauses.",
     "Neutral seat or foreign law only if commercially necessary and the advocate has confirmed enforceability."),
    ("NDA;Services agreement;Vendor agreement;SaaS / technology agreement;Consulting agreement;Employment agreement",
     "Confidentiality",
     [r"confidential", r"non.?disclosure"], [],
     1, "Review",
     "Mutual; defined confidential information with standard exclusions; defined term; return or deletion on request; permitted disclosures.",
     "One-way protection where only one party discloses, with a reasonable survival period."),
    ("Vendor agreement;Services agreement;SaaS / technology agreement;Consulting agreement;Distribution agreement",
     "Payment terms",
     [r"payment terms", r"invoice", r"late fee", r"interest.{0,30}overdue", r"within \d+ days"],
     [r"\b(?:2\d|[3-9]\d)\s?%\s*(?:per annum|p\.a\.|per year)"],
     1, "Review",
     "Clear due dates and invoicing steps; dispute mechanism; tax treatment stated; proportionate late-payment interest.",
     "Shorter payment cycle only with matching service-level commitments."),
    ("Vendor agreement;Services agreement;SaaS / technology agreement;Lease / licence",
     "Automatic renewal",
     [r"auto.?renew", r"automatic(?:ally)? renew", r"renews? automatically", r"automatically (?:be )?(?:renewed|extended)", r"renewal term"],
     [r"auto.?renew", r"automatic(?:ally)? renew", r"renews? automatically", r"automatically (?:be )?(?:renewed|extended)"],
     0, "Review",
     "No automatic renewal, or renewal only with a long notice window and a calendar reminder.",
     "Automatic renewal accepted with a short renewal term and a clear opt-out right."),
    ("Services agreement;SaaS / technology agreement;Consulting agreement;Employment agreement;Vendor agreement",
     "Intellectual property",
     [r"intellectual property", r"work product", r"ownership of deliverables", r"assign.{0,30}copyright"], [],
     1, "Review",
     "Background IP stays with its owner; deliverable ownership or licence matches payment; improvements and feedback addressed.",
     "Broad licence in place of assignment where the client does not need exclusive ownership."),
    ("Vendor agreement;Services agreement;SaaS / technology agreement;Consulting agreement",
     "Data protection and security",
     [r"personal data", r"data protection", r"data processor", r"information security", r"security incident"], [],
     1, "Review",
     "Defined roles and permitted processing; security safeguards; incident notice; sub-processor control; return or deletion. Advocate to confirm alignment with applicable Indian data protection law.",
     "Annexed data processing terms where the main agreement is short."),
    ("Vendor agreement;Services agreement;SaaS / technology agreement;Distribution agreement;Lease / licence",
     "Force majeure",
     [r"force majeure", r"act of god"], [],
     0, "Review",
     "Defined events; notice and mitigation; suspension and a long-stop termination right; payment obligations addressed.",
     "Generic clause accepted if the notice and long-stop termination points are covered elsewhere."),
    ("Employment agreement;Consulting agreement;Distribution agreement",
     "Restrictive covenants",
     [r"non.?compete", r"restraint of trade", r"not engage in any competing"],
     [r"non.?compete", r"restraint of trade", r"not engage in any competing"],
     0, "Review",
     "Avoid broad post-termination restraints; rely on confidentiality and narrowly drawn non-solicit terms. Advocate to advise on enforceability under Indian law.",
     "Limited restraint only during the term of the relationship."),
    ("All", "Assignment and subcontracting",
     [r"assign.{0,30}agreement", r"subcontract", r"sub.?contract"], [],
     0, "Review",
     "Consent required for assignment and subcontracting; responsibility for subcontractors stays with the contracting party.",
     "Permit assignment to group companies or on a merger, with notice."),
]


def init_db():
    with db() as conn:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.executescript(
            """
        CREATE TABLE IF NOT EXISTS matters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_name TEXT NOT NULL,
            matter_title TEXT NOT NULL,
            contract_type TEXT,
            jurisdiction TEXT DEFAULT 'India',
            filename TEXT,
            file_hash TEXT,
            extracted_text TEXT,
            created_at TEXT NOT NULL,
            status TEXT DEFAULT 'Awaiting advocate review',
            analysis_mode TEXT DEFAULT 'Rules',
            lawyer_name TEXT,
            lawyer_notes TEXT,
            lawyer_decision TEXT DEFAULT 'Pending',
            reviewed_at TEXT
        );
        CREATE TABLE IF NOT EXISTS findings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            matter_id INTEGER NOT NULL,
            clause_name TEXT NOT NULL,
            risk_level TEXT NOT NULL,
            source_excerpt TEXT,
            reason TEXT,
            suggested_action TEXT,
            source_location TEXT DEFAULT 'Extracted document text',
            confidence TEXT DEFAULT 'Rule match',
            FOREIGN KEY(matter_id) REFERENCES matters(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS activity (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            matter_id INTEGER,
            event TEXT NOT NULL,
            detail TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY(matter_id) REFERENCES matters(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS drafts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            matter_id INTEGER NOT NULL,
            kind TEXT NOT NULL,
            question TEXT,
            content TEXT NOT NULL,
            quotes_json TEXT,
            gaps_json TEXT,
            truncated INTEGER DEFAULT 0,
            status TEXT DEFAULT 'AI draft — not reviewed',
            reviewed_by TEXT,
            reviewed_at TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY(matter_id) REFERENCES matters(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            full_name TEXT NOT NULL,
            role TEXT NOT NULL,
            enrolment_no TEXT,
            pw_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            iterations INTEGER NOT NULL,
            active INTEGER DEFAULT 1,
            failed_attempts INTEGER DEFAULT 0,
            locked_until TEXT,
            must_change_pw INTEGER DEFAULT 0,
            created_at TEXT NOT NULL,
            last_login TEXT
        );
        CREATE TABLE IF NOT EXISTS playbook (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            applies_to TEXT NOT NULL DEFAULT 'All',
            clause_name TEXT NOT NULL,
            detect_patterns TEXT,
            flag_patterns TEXT,
            required INTEGER DEFAULT 0,
            risk_if_missing TEXT DEFAULT 'Review',
            preferred_position TEXT,
            fallback_position TEXT,
            active INTEGER DEFAULT 1,
            approved_by TEXT,
            approved_at TEXT,
            updated_by TEXT,
            updated_at TEXT
        );
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT NOT NULL,
            username TEXT,
            action TEXT NOT NULL,
            matter_id INTEGER,
            detail TEXT
        );
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        );
        """
        )
        # Safe upgrades for databases created by earlier versions
        ensure_columns(conn, "matters", {
            "jurisdiction": "TEXT DEFAULT 'India'",
            "analysis_mode": "TEXT DEFAULT 'Rules'",
            "lawyer_name": "TEXT",
            "lawyer_notes": "TEXT",
            "lawyer_decision": "TEXT DEFAULT 'Pending'",
            "reviewed_at": "TEXT",
            "created_by": "TEXT",
            "ocr_used": "INTEGER DEFAULT 0",
            "redacted": "INTEGER DEFAULT 0",
            "ai_summary": "TEXT",
            "reviewer_username": "TEXT",
            "reviewer_enrolment": "TEXT",
        })
        ensure_columns(conn, "findings", {
            "source_location": "TEXT DEFAULT 'Extracted document text'",
            "confidence": "TEXT DEFAULT 'Rule match'",
        })
        ensure_columns(conn, "activity", {"username": "TEXT"})
        ensure_columns(conn, "drafts", {"created_by": "TEXT"})
        if conn.execute("SELECT COUNT(*) FROM playbook").fetchone()[0] == 0:
            for types, clause, detect, flags, req, risk, pref, fall in PLAYBOOK_DEFAULTS:
                conn.execute(
                    """INSERT INTO playbook(applies_to,clause_name,detect_patterns,flag_patterns,required,
                       risk_if_missing,preferred_position,fallback_position,active,updated_by,updated_at)
                       VALUES(?,?,?,?,?,?,?,?,1,'starter set',?)""",
                    (types, clause, "\n".join(detect), "\n".join(flags), req, risk, pref, fall, now()),
                )
        for key, val in {"allow_external_ai": "0", "redact_default": "1"}.items():
            conn.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (key, val))


init_db()


def get_setting(key, default=""):
    with db() as conn:
        r = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return r["value"] if r else default


def set_setting(key, value):
    with db() as conn:
        conn.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (key, str(value)))


def acting_user():
    u = st.session_state.get("user")
    return u["username"] if u else "system"


def audit(action, detail="", matter_id=None):
    with db() as conn:
        conn.execute("INSERT INTO audit_log(ts,username,action,matter_id,detail) VALUES(?,?,?,?,?)",
                     (now(), acting_user(), action, matter_id, detail))


def log_event(mid, event, detail=""):
    with db() as conn:
        conn.execute("INSERT INTO activity(matter_id,event,detail,created_at,username) VALUES(?,?,?,?,?)",
                     (mid, event, detail, now(), acting_user()))


def backup_bytes():
    """Consistent copy of the SQLite database (safe even while the app is running)."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".db")
    tmp.close()
    src = sqlite3.connect(DB_PATH)
    dst = sqlite3.connect(tmp.name)
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()
    try:
        return Path(tmp.name).read_bytes()
    finally:
        os.unlink(tmp.name)


# ================================================================ ACCOUNTS
def hash_pw(password, salt=None, iterations=PBKDF2_ITER):
    salt = salt or secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt), iterations)
    return dk.hex(), salt, iterations


def verify_pw(password, user_row):
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(user_row["salt"]),
                             user_row["iterations"])
    return hmac.compare_digest(dk.hex(), user_row["pw_hash"])


def validate_password(pw):
    if len(pw) < MIN_PW_LEN:
        raise ValueError(f"Password must be at least {MIN_PW_LEN} characters.")
    if not (re.search(r"[A-Za-z]", pw) and re.search(r"\d", pw)):
        raise ValueError("Password must contain both letters and numbers.")


def user_count():
    with db() as conn:
        return conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]


def get_user_by_name(username):
    with db() as conn:
        return conn.execute("SELECT * FROM users WHERE username=?", ((username or "").strip().lower(),)).fetchone()


def get_user_by_id(uid):
    with db() as conn:
        return conn.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()


def list_users():
    with db() as conn:
        return conn.execute("SELECT * FROM users ORDER BY id").fetchall()


def create_user(username, full_name, role, password, enrolment="", must_change=False):
    username = (username or "").strip().lower()
    if not re.fullmatch(r"[a-z0-9._-]{3,40}", username):
        raise ValueError("Username must be 3–40 characters: lowercase letters, numbers, dot, dash or underscore.")
    if not (full_name or "").strip():
        raise ValueError("Enter the person's full name.")
    if role not in ROLES:
        raise ValueError("Choose a valid role.")
    validate_password(password)
    h, salt, it = hash_pw(password)
    try:
        with db() as conn:
            conn.execute(
                """INSERT INTO users(username,full_name,role,enrolment_no,pw_hash,salt,iterations,must_change_pw,created_at)
                   VALUES(?,?,?,?,?,?,?,?,?)""",
                (username, full_name.strip(), role, (enrolment or "").strip(), h, salt, it, int(must_change), now()),
            )
    except sqlite3.IntegrityError:
        raise ValueError("That username already exists.")


def set_password(uid, new_password, must_change=False):
    validate_password(new_password)
    h, salt, it = hash_pw(new_password)
    with db() as conn:
        conn.execute("UPDATE users SET pw_hash=?,salt=?,iterations=?,must_change_pw=?,failed_attempts=0,locked_until=NULL WHERE id=?",
                     (h, salt, it, int(must_change), uid))


def authenticate(username, password):
    generic = "Incorrect username or password."
    u = get_user_by_name(username)
    if not u or not u["active"]:
        hash_pw(password or "x")  # keep timing similar for unknown users
        return None, generic
    if u["locked_until"] and now() < u["locked_until"]:
        return None, f"Too many failed attempts. Try again after {u['locked_until'][11:16]}."
    if verify_pw(password or "", u):
        with db() as conn:
            conn.execute("UPDATE users SET failed_attempts=0,locked_until=NULL,last_login=? WHERE id=?", (now(), u["id"]))
        return get_user_by_id(u["id"]), ""
    fails = (u["failed_attempts"] or 0) + 1
    locked = None
    if fails >= MAX_FAILED_LOGINS:
        locked = (datetime.now() + timedelta(minutes=LOCK_MINUTES)).isoformat(timespec="seconds")
        fails = 0
    with db() as conn:
        conn.execute("UPDATE users SET failed_attempts=?,locked_until=? WHERE id=?", (fails, locked, u["id"]))
    return None, generic if not locked else f"Too many failed attempts. Account locked for {LOCK_MINUTES} minutes."


def current_user():
    sess = st.session_state.get("user")
    if not sess:
        return None
    if time.time() - st.session_state.get("last_active", 0) > SESSION_TIMEOUT_MIN * 60:
        st.session_state.clear()
        st.session_state["timed_out"] = True
        return None
    row = get_user_by_id(sess["id"])
    if not row or not row["active"]:
        st.session_state.clear()
        return None
    st.session_state["last_active"] = time.time()
    st.session_state["user"] = dict(row)
    return dict(row)


def can(action):
    u = st.session_state.get("user")
    return bool(u) and u["role"] in PERMS.get(action, set())


def render_login():
    _, mid, _ = st.columns([1, 1.3, 1])
    with mid:
        with st.container(border=True):
            if user_count() == 0:
                st.markdown("### Create the administrator account")
                st.caption("This is the first launch. The administrator can add advocates and staff afterwards.")
                with st.form("setup_form"):
                    full = st.text_input("Full name")
                    uname = st.text_input("Username", placeholder="e.g. priya.rao")
                    pw1 = st.text_input("Password", type="password", help=f"At least {MIN_PW_LEN} characters, letters and numbers.")
                    pw2 = st.text_input("Confirm password", type="password")
                    go = st.form_submit_button("Create administrator", type="primary", use_container_width=True)
                if go:
                    if pw1 != pw2:
                        st.error("The passwords do not match.")
                    else:
                        try:
                            create_user(uname, full, "admin", pw1)
                            audit("Administrator created", uname)
                            st.success("Administrator created. Sign in below.")
                            st.rerun()
                        except ValueError as e:
                            st.error(str(e))
            else:
                st.markdown("### Sign in")
                if st.session_state.pop("timed_out", False):
                    st.info("You were signed out after a period of inactivity.")
                with st.form("login_form"):
                    uname = st.text_input("Username")
                    pw = st.text_input("Password", type="password")
                    go = st.form_submit_button("Sign in", type="primary", use_container_width=True)
                if go:
                    u, msg = authenticate(uname, pw)
                    if u:
                        st.session_state["user"] = dict(u)
                        st.session_state["last_active"] = time.time()
                        audit("Sign in")
                        st.rerun()
                    else:
                        with db() as conn:
                            conn.execute("INSERT INTO audit_log(ts,username,action,detail) VALUES(?,?,?,?)",
                                         (now(), (uname or "")[:40].lower(), "Failed sign in", ""))
                        st.error(msg)
            st.caption("Prototype accounts protect access to this app. They are not a substitute for production security.")


# ======================================================== DOCUMENT PROCESSING
@st.cache_resource
def ocr_ready():
    try:
        import pytesseract

        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def ocr_page(page):
    import pytesseract
    from PIL import Image

    pix = page.get_pixmap(dpi=200)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    return pytesseract.image_to_string(img, lang=os.getenv("LEXPILOT_OCR_LANG", "eng"))


def extract_pdf(data, use_ocr=False):
    info = {"pages": 0, "ocr_pages": 0, "blank_pages": 0}
    parts = []
    with fitz.open(stream=data, filetype="pdf") as doc:
        if doc.needs_pass:
            raise ValueError("This PDF is password-protected. Remove the password and upload it again.")
        if doc.page_count > MAX_PAGES:
            raise ValueError(f"This PDF has {doc.page_count} pages; the limit is {MAX_PAGES}.")
        info["pages"] = doc.page_count
        for i, page in enumerate(doc, start=1):
            txt = page.get_text("text").strip()
            marker = f"[PAGE {i}]"
            if len(txt) < 25 and use_ocr and info["ocr_pages"] < MAX_OCR_PAGES:
                try:
                    ocr_txt = ocr_page(page).strip()
                except Exception:
                    ocr_txt = ""
                if ocr_txt:
                    txt, marker = ocr_txt, f"[PAGE {i} — OCR]"
                    info["ocr_pages"] += 1
            if len(txt) < 25:
                info["blank_pages"] += 1
            if txt:
                parts.append(f"{marker}\n{txt}")
    return "\n\n".join(parts), info


def extract_docx(data):
    document = Document(io.BytesIO(data))
    parts = []
    for i, p in enumerate(document.paragraphs, start=1):
        if p.text.strip():
            parts.append(f"[PARAGRAPH {i}] {p.text}")
    for ti, table in enumerate(document.tables, start=1):
        for ri, row in enumerate(table.rows, start=1):
            parts.append(f"[TABLE {ti} ROW {ri}] " + " | ".join(c.text for c in row.cells))
    return "\n".join(parts)


def extract_txt(data):
    for enc in ("utf-8", "utf-16", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="ignore")


def clean_text(text):
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


REDACTIONS = [
    (r"\b\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]\b", "[GSTIN]"),
    (r"\b[A-Z]{5}\d{4}[A-Z]\b", "[PAN]"),
    (r"(?<!\d)\d{4}\s?\d{4}\s?\d{4}(?!\d)", "[AADHAAR]"),
    (r"[\w.+-]+@[\w-]+\.[\w.-]+", "[EMAIL]"),
    (r"(?<!\d)(?:\+91[\s-]?)?[6-9]\d{9}(?!\d)", "[PHONE]"),
]


def redact_text(text):
    """Best-effort masking of obvious identifiers. Names and addresses are NOT removed."""
    for pat, label in REDACTIONS:
        text = re.sub(pat, label, text)
    return text


# ============================================================ RULES + EXTRACTION
RULES = [
    ("Limitation of liability",
     [r"unlimited liability", r"without limitation", r"no limit.{0,30}liabilit", r"liabilit.{0,30}unlimited"],
     "High",
     "The wording may create uncapped or unusually broad liability exposure.",
     "Ask the advocate to assess a commercially appropriate cap, exclusions and negotiated carve-outs."),
    ("Indemnity", [r"indemnif", r"hold harmless"], "Review",
     "Indemnity obligations can create significant financial exposure.",
     "Review covered losses, third-party claims, defence control, notice, mitigation and interaction with liability caps."),
    ("Termination", [r"terminat"], "Review",
     "Termination rights and consequences may be asymmetric or incomplete.",
     "Check mutuality, notice, cure periods, payment on termination, transition assistance and survival."),
    ("Automatic renewal", [r"automatic renewal", r"auto.?renew", r"renews? automatically", r"automatically (?:be )?(?:renewed|extended)", r"renewal term"], "Review",
     "An automatic renewal may create an unintended extension or notice deadline.",
     "Confirm renewal duration, notice window, notice method and calendar reminders."),
    ("Payment and late fees",
     [r"payment terms", r"invoice", r"late fee", r"interest.{0,30}overdue", r"within \d+ days"], "Review",
     "Payment wording affects cash flow and dispute risk.",
     "Confirm due dates, invoice requirements, dispute process, taxes, withholding and late-payment calculation."),
    ("Confidentiality", [r"confidential", r"non.?disclosure"], "Review",
     "Confidentiality language should define protected information, permitted use and exceptions.",
     "Check exclusions, permitted disclosures, duration, return/deletion and security obligations."),
    ("Intellectual property",
     [r"intellectual property", r"work product", r"ownership of deliverables", r"assign.{0,30}copyright"], "Review",
     "Ownership and licensing of pre-existing IP, deliverables and improvements may be unclear.",
     "Separate background IP from project deliverables; confirm assignment/licence scope and payment conditions."),
    ("Dispute resolution and governing law",
     [r"arbitrat", r"governing law", r"jurisdiction", r"courts at"], "Review",
     "Dispute forum, governing law, seat/venue and escalation procedures should be consistent.",
     "Verify governing law, exclusive jurisdiction, arbitration seat, venue and appointment procedure."),
    ("Data protection and security",
     [r"personal data", r"data protection", r"data processor", r"information security", r"security incident"], "Review",
     "The contract may involve personal data or security duties needing allocation.",
     "Review roles, permitted processing, security controls, incident notice, subprocessors, retention and deletion."),
    ("Assignment and subcontracting", [r"assign.{0,30}agreement", r"subcontract", r"sub.?contract"], "Review",
     "Transfer or subcontracting rights can change who performs contractual obligations.",
     "Check consent requirements, permitted group transfers and responsibility for subcontractors."),
    ("Force majeure", [r"force majeure", r"act of god"], "Review",
     "The clause may affect performance during events outside a party's control.",
     "Check event definition, notice, mitigation, suspension period and termination rights."),
    ("Non-compete / restrictive covenant",
     [r"non.?compete", r"restraint of trade", r"not engage in any competing"], "Review",
     "Restrictive covenants can raise enforceability and commercial scope questions.",
     "Obtain advocate assessment of scope, duration, geography, parties and applicable Indian law."),
]
RULE_STEMS = {
    "Limitation of liability": r"liabil", "Indemnity": r"indemn", "Termination": r"termin",
    "Automatic renewal": r"renew", "Payment and late fees": r"payment|fee|invoice",
    "Confidentiality": r"confiden|disclos", "Intellectual property": r"intellectual|\bip\b|ownership",
    "Dispute resolution and governing law": r"govern|arbitr|dispute|jurisdiction",
    "Data protection and security": r"data|privacy|security", "Assignment and subcontracting": r"assign|subcontract",
    "Force majeure": r"force", "Non-compete / restrictive covenant": r"compete|restrict|restraint",
}


def snippets_for(text, patterns, limit=4):
    lines = re.split(r"(?<=[.;])\s+|\n+", text)
    hits = []
    for line in lines:
        line = line.strip()
        if len(line) > 22 and any(re.search(p, line, re.I) for p in patterns):
            hits.append(line[:1800])
    return hits[:limit]


def rule_analysis(text):
    findings = []
    for name, patterns, risk, reason, action in RULES:
        hits = snippets_for(text, patterns)
        if hits:
            findings.append({
                "name": name, "risk": risk, "excerpt": "\n---\n".join(hits),
                "reason": reason, "action": action,
                "location": "See page/paragraph markers in matched excerpt",
                "confidence": "Keyword match — requires human validation",
            })
    if not findings:
        findings.append({
            "name": "No configured rule matched", "risk": "Unassessed", "excerpt": "",
            "reason": "The configured keyword rules did not identify a matching passage. "
                      "This is not evidence that the contract is safe or complete.",
            "action": "Review the full contract, transaction context and applicable law.",
            "location": "Whole document", "confidence": "No match — not a safety score",
        })
    return findings


def combine_findings(ai_findings, rule_findings):
    """Keep AI findings, then add rule hits the AI did not mention (safety net)."""
    names = " ".join(f["name"].lower() for f in ai_findings)
    out = list(ai_findings)
    for rf in rule_findings:
        if rf["risk"] == "Unassessed":
            continue
        stem = RULE_STEMS.get(rf["name"])
        if stem and re.search(stem, names):
            continue
        rf = dict(rf)
        rf["confidence"] = "Rules safety-net (the AI did not flag this) — keyword match, requires validation"
        out.append(rf)
    return out


def valid_patterns(text):
    pats = [p.strip() for p in (text or "").splitlines() if p.strip()]
    if len(pats) > 12:
        raise ValueError("Use at most 12 patterns per field.")
    for p in pats:
        if len(p) > 200:
            raise ValueError("Each pattern must be 200 characters or fewer.")
        try:
            re.compile(p)
        except re.error as e:
            raise ValueError(f"Invalid pattern “{p[:40]}”: {e}")
    return pats


def applicable(item, contract_type):
    types = [t.strip() for t in (item["applies_to"] or "All").split(";") if t.strip()]
    return "All" in types or contract_type in types


def run_playbook(text, contract_type, items):
    results = []
    for it in items:
        if not it["active"] or not applicable(it, contract_type):
            continue
        detect = [p for p in (it["detect_patterns"] or "").splitlines() if p.strip()]
        flags = [p for p in (it["flag_patterns"] or "").splitlines() if p.strip()]
        try:
            found = snippets_for(text, detect, 3) if detect else []
            flagged = snippets_for(text, flags, 3) if flags else []
        except re.error:
            found, flagged = [], []
        if flagged:
            status, icon = "Red-flag wording found", "🔴"
        elif found:
            status, icon = "Located — compare with the preferred position", "🟢"
        elif it["required"]:
            status, icon = "Not located — check the full document", "🟠"
        else:
            status, icon = "Not located (optional clause)", "⚪"
        results.append({
            "clause": it["clause_name"], "status": status, "icon": icon, "required": bool(it["required"]),
            "risk_if_missing": it["risk_if_missing"], "found": found, "flagged": flagged,
            "preferred": it["preferred_position"] or "", "fallback": it["fallback_position"] or "",
            "approved": bool(it["approved_by"]), "approved_by": it["approved_by"],
        })
    return results


MONTHS = "January|February|March|April|May|June|July|August|September|October|November|December"
PERIOD_RE = re.compile(r"\b(\d{1,3})(?:\s*\([^)]{1,30}\))?\s*(business days|working days|calendar days|days|weeks|months|years)\b", re.I)
DATE_RE = re.compile(
    rf"\b(?:\d{{1,2}}[/-]\d{{1,2}}[/-]\d{{2,4}}|\d{{1,2}}(?:st|nd|rd|th)?\s+(?:{MONTHS})\s+\d{{4}}|(?:{MONTHS})\s+\d{{1,2}},?\s+\d{{4}})\b",
    re.I)
MONEY_RE = re.compile(r"(?:₹|\bRs\.?|\bINR|\bUSD|US\$|\$)\s?\d[\d,]*(?:\.\d+)?(?:\s?(?:lakhs?|crores?|million|billion|thousand))?", re.I)
PERCENT_RE = re.compile(r"\b\d{1,3}(?:\.\d+)?\s?%")


def key_terms(text):
    flat = re.sub(r"\[(?:PAGE|PARAGRAPH|TABLE)[^\]]*\]", " ", text)
    flat = re.sub(r"\s+", " ", flat)
    out, seen = [], set()
    for kind, rx in [("Time period", PERIOD_RE), ("Date", DATE_RE), ("Amount", MONEY_RE), ("Percentage", PERCENT_RE)]:
        count = 0
        for m in rx.finditer(flat):
            ctx = flat[max(0, m.start() - 110): m.end() + 110].strip()
            key = (kind, m.group(0).lower(), ctx[:60])
            if key in seen:
                continue
            seen.add(key)
            out.append({"Type": kind, "Value": m.group(0).strip(), "Context": "…" + ctx + "…"})
            count += 1
            if count >= 25:
                break
    return out


def sentences(text):
    flat = re.sub(r"\[(?:PAGE|PARAGRAPH|TABLE)[^\]]*\]", " ", text or "")
    flat = re.sub(r"\s+", " ", flat).strip()
    return [s.strip() for s in re.split(r"(?<=[.;:])\s+", flat) if s.strip()]


def diff_texts(a, b):
    return list(difflib.unified_diff(sentences(a), sentences(b), "Version A", "Version B", lineterm="", n=0))


# ================================================================ GEMINI AI
SYSTEM_PROMPT = """You assist a qualified Indian advocate by preparing preliminary contract-review issue spotting.
You are not a lawyer and must not provide a final legal opinion, declare enforceability, or claim a contract is safe.
Use only the supplied contract text. Do not invent clauses, facts, law, page numbers, parties, dates or citations.
Identify concrete passages that may deserve advocate review. If a clause is absent, do not call it legally defective; say it was not located in the supplied text and that the full document must be checked.
Return valid JSON only with this schema:
{
 "summary": "short neutral summary",
 "findings": [
  {"clause_name":"...", "risk_level":"High|Review|Informational|Unassessed",
   "source_excerpt":"verbatim short excerpt from supplied text",
   "reason":"why this wording may matter, cautiously phrased",
   "suggested_action":"questions/checks for advocate review",
   "source_location":"page/paragraph marker if present, otherwise Not identified",
   "confidence":"High|Medium|Low and brief basis"}
 ]
}
Maximum 15 findings. Distinguish text evidence from assumptions. Do not draft final advice."""

DRAFT_KINDS = {
    "Ask a question": "Answer the advocate's question using only the contract text.",
    "Plain-English summary": "Write a neutral plain-English summary: purpose, parties (if stated), term, payment, key obligations and notable areas that may need advocate attention.",
    "Negotiation checklist": "Prepare a checklist of points the advocate may wish to raise or verify with the client, each tied to specific contract wording.",
    "Client cover note (draft)": "Draft a short cover note to the client summarising the preliminary review points. Neutral tone. No legal conclusions or assurances. State that advocate review is pending.",
    "Email to counterparty (draft)": "Draft a short, polite email to the counterparty asking clarification questions about specific clauses. No admissions, commitments or legal conclusions.",
}

ASSIST_PROMPT = """You assist a qualified Indian advocate by preparing DRAFT working material from one contract.
You are not a lawyer. Do not give a final legal opinion, declare enforceability or legality, or say a contract is safe.
Use ONLY the supplied contract text. Do not invent clauses, parties, dates, figures, law or citations.
If the text does not answer something, say so in "gaps" instead of guessing.
Every factual statement about the contract must be supported by a verbatim excerpt in "quotes" (copied exactly, max 40 words each).
Write "content" in Markdown. Return valid JSON only:
{"content": "markdown text", "quotes": ["verbatim excerpt", "..."], "gaps": ["what was not found in the text"]}"""

GEMINI_MODELS = ["gemini-2.5-flash", "gemini-2.5-flash-lite"]  # tried in order
TRANSIENT_MARKERS = ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "500", "INTERNAL", "504", "DEADLINE")


def get_api_key():
    try:
        key = st.secrets.get("GEMINI_API_KEY", "")
        if key:
            return key
    except Exception:
        pass
    return os.getenv("GEMINI_API_KEY", "")


def ai_gate():
    """Central switch: external AI only if an admin enabled it, a key exists and the user is within limits."""
    if get_setting("allow_external_ai", "0") != "1":
        return False, "External AI is switched off. An administrator can enable it under Admin → AI & privacy."
    if not get_api_key():
        return False, "Gemini API key not configured. Add GEMINI_API_KEY to Streamlit secrets or the environment."
    since = (datetime.now() - timedelta(hours=1)).isoformat(timespec="seconds")
    with db() as conn:
        n = conn.execute("SELECT COUNT(*) FROM audit_log WHERE action='AI call' AND username=? AND ts>=?",
                         (acting_user(), since)).fetchone()[0]
    if n >= AI_HOURLY_LIMIT:
        return False, f"Hourly AI limit of {AI_HOURLY_LIMIT} requests reached. Try again later."
    return True, ""


def generate_json(client, prompt, temperature):
    """Call Gemini in JSON mode; retry busy/transient errors, then fall back to the next model."""
    last_error = None
    for model in GEMINI_MODELS:
        for attempt in range(3):
            try:
                return client.models.generate_content(
                    model=model, contents=prompt,
                    config={"response_mime_type": "application/json", "temperature": temperature},
                )
            except Exception as e:
                last_error = e
                if not any(m in str(e) for m in TRANSIENT_MARKERS):
                    raise
                if attempt < 2:
                    time.sleep(2 * (attempt + 1))
    raise last_error


def norm(s):
    s = re.sub(r"\[(?:PAGE|PARAGRAPH|TABLE)[^\]]*\]", " ", s or "")
    return re.sub(r"\s+", " ", s).strip().lower()


def verify_quotes(quotes, text):
    haystack = norm(text)
    out = []
    for q in quotes:
        if isinstance(q, str) and norm(q):
            out.append({"quote": q.strip()[:600], "found": norm(q) in haystack})
    return out


def parse_json(raw):
    raw = (raw or "{}").strip()
    raw = re.sub(r"^```(?:json)?|```$", "", raw, flags=re.M).strip()
    data = json.loads(raw)
    return data if isinstance(data, dict) else {}


def gemini_analysis(text, redact=True, matter_ref=None):
    ok, msg = ai_gate()
    if not ok:
        return None, msg
    try:
        from google import genai

        sent = redact_text(text) if redact else text
        client = genai.Client(api_key=get_api_key(), http_options={"api_version": "v1beta"})
        prompt = SYSTEM_PROMPT + "\n\nCONTRACT TEXT:\n" + sent[:MAX_CHARS]
        data = parse_json(generate_json(client, prompt, 0.1).text)
        audit("AI call", f"issue spotting; redacted={redact}; chars={min(len(sent), MAX_CHARS)}", matter_ref)
        findings = []
        for f in list(data.get("findings", []))[:15]:
            risk = f.get("risk_level", "Review")
            excerpt = str(f.get("source_excerpt", ""))[:2200]
            basis = str(f.get("confidence", "Not stated"))[:160]
            if excerpt:
                basis += " | excerpt verified in text" if norm(excerpt) in norm(sent) else " | EXCERPT NOT FOUND IN TEXT — verify"
            findings.append({
                "name": str(f.get("clause_name", "Unspecified issue"))[:160],
                "risk": risk if risk in RISK_LEVELS else "Review",
                "excerpt": excerpt,
                "reason": str(f.get("reason", ""))[:1200],
                "action": str(f.get("suggested_action", ""))[:1200],
                "location": str(f.get("source_location", "Not identified"))[:200],
                "confidence": basis[:300],
            })
        return {"summary": str(data.get("summary", ""))[:3000], "findings": findings,
                "truncated": len(sent) > MAX_CHARS}, None
    except Exception as e:
        return None, f"AI analysis failed: {type(e).__name__}: {e}"


def gemini_assistant(text, kind, question="", redact=True, matter_ref=None):
    ok, msg = ai_gate()
    if not ok:
        return None, msg
    try:
        from google import genai

        sent = redact_text(text) if redact else text
        client = genai.Client(api_key=get_api_key(), http_options={"api_version": "v1beta"})
        task = DRAFT_KINDS[kind]
        if kind == "Ask a question":
            task += "\n\nQUESTION: " + question.strip()[:2000]
        prompt = f"{ASSIST_PROMPT}\n\nTASK: {task}\n\nCONTRACT TEXT:\n{sent[:MAX_CHARS]}"
        data = parse_json(generate_json(client, prompt, 0.2).text)
        audit("AI call", f"{kind}; redacted={redact}; chars={min(len(sent), MAX_CHARS)}", matter_ref)
        content = str(data.get("content", "")).strip()[:12000]
        if not content:
            return None, "The AI returned an empty draft. Try rephrasing or run it again."
        return {
            "content": content,
            "quotes": verify_quotes(list(data.get("quotes", []))[:15], sent),
            "gaps": [str(g)[:500] for g in list(data.get("gaps", []))[:10]],
            "truncated": len(sent) > MAX_CHARS,
        }, None
    except Exception as e:
        return None, f"AI draft failed: {type(e).__name__}: {e}"


# ================================================================== STORAGE
def create_matter(meta, extracted, findings):
    with db() as conn:
        cur = conn.execute(
            """INSERT INTO matters
            (client_name,matter_title,contract_type,jurisdiction,filename,file_hash,extracted_text,created_at,
             analysis_mode,status,created_by,ocr_used,redacted,ai_summary)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (meta["client"], meta["title"], meta["ctype"], meta["jurisdiction"], meta["filename"], meta["digest"],
             extracted, now(), meta["mode"], "Awaiting advocate review", acting_user(),
             int(meta.get("ocr_used", 0)), int(meta.get("redacted", 0)), meta.get("summary") or None),
        )
        mid = cur.lastrowid
        for f in findings:
            conn.execute(
                """INSERT INTO findings
                (matter_id,clause_name,risk_level,source_excerpt,reason,suggested_action,source_location,confidence)
                VALUES(?,?,?,?,?,?,?,?)""",
                (mid, f["name"], f["risk"], f["excerpt"], f["reason"], f["action"], f["location"], f["confidence"]),
            )
    log_event(mid, "Matter created", f"Analysis mode: {meta['mode']}; file: {meta['filename']}")
    audit("Matter created", meta["filename"], mid)
    return mid


def list_matters():
    with db() as conn:
        return conn.execute("SELECT * FROM matters ORDER BY id DESC").fetchall()


def find_duplicate(digest):
    with db() as conn:
        return conn.execute("SELECT id, matter_title FROM matters WHERE file_hash=? ORDER BY id DESC LIMIT 1",
                            (digest,)).fetchone()


def load_matter(mid):
    with db() as conn:
        matter = conn.execute("SELECT * FROM matters WHERE id=?", (mid,)).fetchone()
        findings = conn.execute("SELECT * FROM findings WHERE matter_id=? ORDER BY id", (mid,)).fetchall()
        events = conn.execute("SELECT * FROM activity WHERE matter_id=? ORDER BY id DESC", (mid,)).fetchall()
    return matter, findings, events


def save_review(mid, user, notes, decision):
    if user["role"] not in PERMS["review"]:
        raise PermissionError("Only advocates and administrators can record review decisions.")
    status = {APPROVED: "Advocate review complete",
              "Review in progress": "Advocate review in progress"}.get(decision, "Awaiting advocate review")
    with db() as conn:
        conn.execute(
            """UPDATE matters SET lawyer_name=?,lawyer_notes=?,lawyer_decision=?,reviewed_at=?,status=?,
               reviewer_username=?,reviewer_enrolment=? WHERE id=?""",
            (user["full_name"], notes, decision, now(), status, user["username"], user.get("enrolment_no") or "", mid),
        )
    log_event(mid, "Advocate review updated", decision)
    audit("Review updated", decision, mid)


def delete_matter(mid):
    with db() as conn:
        ref = conn.execute("SELECT matter_title, filename FROM matters WHERE id=?", (mid,)).fetchone()
        conn.execute("DELETE FROM matters WHERE id=?", (mid,))
    audit("Matter deleted", f"{ref['matter_title']} / {ref['filename']}" if ref else "", mid)


def save_draft(mid, kind, question, res):
    with db() as conn:
        cur = conn.execute(
            """INSERT INTO drafts(matter_id,kind,question,content,quotes_json,gaps_json,truncated,status,created_at,created_by)
            VALUES(?,?,?,?,?,?,?,?,?,?)""",
            (mid, kind, question, res["content"], json.dumps(res["quotes"]), json.dumps(res["gaps"]),
             int(res["truncated"]), DRAFT_UNREVIEWED, now(), acting_user()),
        )
        did = cur.lastrowid
    log_event(mid, "AI draft saved", f"{kind} (draft #{did}) — not reviewed")
    return did


def list_drafts(mid):
    with db() as conn:
        return conn.execute("SELECT * FROM drafts WHERE matter_id=? ORDER BY id DESC", (mid,)).fetchall()


def review_draft(did, mid, user):
    if user["role"] not in PERMS["review"]:
        raise PermissionError("Only advocates and administrators can mark drafts as reviewed.")
    with db() as conn:
        conn.execute("UPDATE drafts SET status=?, reviewed_by=?, reviewed_at=? WHERE id=?",
                     (DRAFT_REVIEWED, user["full_name"], now(), did))
    log_event(mid, "AI draft marked reviewed", f"Draft #{did}")


def delete_draft(did, mid):
    with db() as conn:
        conn.execute("DELETE FROM drafts WHERE id=?", (did,))
    log_event(mid, "AI draft deleted", f"Draft #{did}")


def list_playbook():
    with db() as conn:
        return conn.execute("SELECT * FROM playbook ORDER BY clause_name, id").fetchall()


def save_playbook_item(item_id, data, approve, user):
    detect = valid_patterns(data["detect_patterns"])
    flags = valid_patterns(data["flag_patterns"])
    if not data["clause_name"].strip():
        raise ValueError("Enter a clause name.")
    if not data["applies_to"]:
        raise ValueError("Choose at least one contract type.")
    approved_by = user["full_name"] if approve and user["role"] in PERMS["playbook_edit"] else None
    approved_at = now() if approved_by else None
    vals = (";".join(data["applies_to"]), data["clause_name"].strip(), "\n".join(detect), "\n".join(flags),
            int(data["required"]), data["risk_if_missing"], data["preferred_position"].strip(),
            data["fallback_position"].strip(), int(data["active"]), approved_by, approved_at, user["username"], now())
    with db() as conn:
        if item_id:
            conn.execute(
                """UPDATE playbook SET applies_to=?,clause_name=?,detect_patterns=?,flag_patterns=?,required=?,
                   risk_if_missing=?,preferred_position=?,fallback_position=?,active=?,approved_by=?,approved_at=?,
                   updated_by=?,updated_at=? WHERE id=?""", vals + (item_id,))
        else:
            conn.execute(
                """INSERT INTO playbook(applies_to,clause_name,detect_patterns,flag_patterns,required,risk_if_missing,
                   preferred_position,fallback_position,active,approved_by,approved_at,updated_by,updated_at)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""", vals)
    audit("Playbook saved", data["clause_name"] + (" (approved)" if approved_by else " (unapproved)"))


def delete_playbook_item(item_id):
    with db() as conn:
        conn.execute("DELETE FROM playbook WHERE id=?", (item_id,))
    audit("Playbook item deleted", str(item_id))


# ================================================================== EXPORTS
def report_markdown(m, findings, pb, terms, prepared_by):
    out = [
        "# LexPilot — Preliminary Contract Review Pack", "",
        "> INTERNAL WORKING DOCUMENT. Not legal advice, not an advocate's opinion and not approved for "
        "client delivery unless the appropriate advocate has independently reviewed and authorised it.", "",
        f"**Client:** {m['client_name']}  ",
        f"**Matter:** {m['matter_title']}  ",
        f"**Contract type:** {m['contract_type']}  ",
        f"**Jurisdiction:** {m['jurisdiction']}  ",
        f"**File:** {m['filename']}  ",
        f"**Analysis method:** {m['analysis_mode']}{' (OCR used on scanned pages)' if m['ocr_used'] else ''}  ",
        f"**Created:** {m['created_at']}  ",
        f"**Pack prepared by:** {prepared_by} on {now()}", "",
        "## Limitations",
        "Automated analysis can be incomplete or wrong. A qualified advocate must examine the complete "
        "agreement, transaction facts, instructions and current applicable law. A missing match does not "
        "mean a clause is absent or a contract is safe.", "",
    ]
    if m["ai_summary"]:
        out += ["## AI summary (unverified)", m["ai_summary"], ""]
    out.append("## Findings")
    for i, f in enumerate(findings, 1):
        out += [
            f"### {i}. {f['clause_name']} — {f['risk_level']}",
            f"**Source location:** {f['source_location']}",
            f"**Confidence / basis:** {f['confidence']}",
            f"**Why it may matter:** {f['reason']}",
            f"**Advocate review action:** {f['suggested_action']}",
            "**Matched source wording:**",
            f"> {(f['source_excerpt'] or 'No excerpt captured.')[:3000]}", "",
        ]
    if pb:
        out += ["## Playbook comparison",
                "Positions marked “starter” are not advocate-approved and are shown for discussion only.", ""]
        for r in pb:
            out += [f"### {r['icon']} {r['clause']} — {r['status']}",
                    f"**Preferred position ({'approved by ' + r['approved_by'] if r['approved'] else 'starter, not approved'}):** {r['preferred']}",
                    f"**Fallback:** {r['fallback']}"]
            for s in (r["flagged"] or r["found"])[:2]:
                out.append(f"> {s[:600]}")
            out.append("")
    if terms:
        out += ["## Key terms detected (keyword extraction — verify against the contract)", "",
                "| Type | Value | Context |", "|---|---|---|"]
        for t in terms[:40]:
            out.append(f"| {t['Type']} | {t['Value']} | {t['Context'].replace('|', '/')[:200]} |")
        out.append("")
    out += [
        "## Advocate review record",
        f"**Reviewer:** {m['lawyer_name'] or 'Not recorded'}"
        + (f" (account: {m['reviewer_username']}; enrolment no. as entered: {m['reviewer_enrolment'] or 'not entered'})" if m["reviewer_username"] else ""),
        f"**Decision:** {m['lawyer_decision']}",
        f"**Notes:** {m['lawyer_notes'] or 'No notes recorded.'}",
    ]
    return "\n".join(out)


def make_docx_pack(m, findings, pb, terms, prepared_by):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.7)
    sec.bottom_margin = Inches(0.7)
    title = doc.add_heading("LexPilot — Preliminary Contract Review Pack", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = doc.add_paragraph().add_run("INTERNAL WORKING DOCUMENT — NOT LEGAL ADVICE")
    r.bold = True
    r.font.color.rgb = RGBColor(170, 72, 55)
    doc.add_paragraph(
        "Automated issue spotting is preliminary and may be incomplete or incorrect. An independent "
        "qualified advocate must review the complete agreement, context and applicable law before any "
        "client-facing legal deliverable."
    )
    doc.add_heading("Matter details", level=1)
    for label, val in [
        ("Client", m["client_name"]), ("Matter", m["matter_title"]),
        ("Contract type", m["contract_type"]), ("Jurisdiction", m["jurisdiction"]),
        ("Source file", m["filename"]),
        ("Analysis method", m["analysis_mode"] + (" (OCR used on scanned pages)" if m["ocr_used"] else "")),
        ("Created", m["created_at"]), ("Pack prepared by", f"{prepared_by} on {now()}"),
    ]:
        doc.add_paragraph(f"{label}: {val or 'Not recorded'}")
    if m["ai_summary"]:
        doc.add_heading("AI summary (unverified)", level=1)
        doc.add_paragraph(m["ai_summary"])
    doc.add_heading("Preliminary findings", level=1)
    for i, f in enumerate(findings, 1):
        doc.add_heading(f"{i}. {f['clause_name']} — {f['risk_level']}", level=2)
        for label, val in [
            ("Source location", f["source_location"]), ("Confidence / basis", f["confidence"]),
            ("Why it may matter", f["reason"]), ("Advocate review action", f["suggested_action"]),
            ("Matched source wording", f["source_excerpt"] or "No excerpt captured."),
        ]:
            p = doc.add_paragraph()
            p.add_run(label + ": ").bold = True
            p.add_run(val or "")
    if pb:
        doc.add_heading("Playbook comparison", level=1)
        doc.add_paragraph("Positions marked “starter” are not advocate-approved and are shown for discussion only.")
        for r_ in pb:
            doc.add_heading(f"{r_['clause']} — {r_['status']}", level=2)
            p = doc.add_paragraph()
            p.add_run("Preferred position (" + (f"approved by {r_['approved_by']}" if r_["approved"] else "starter, not approved") + "): ").bold = True
            p.add_run(r_["preferred"])
            p = doc.add_paragraph()
            p.add_run("Fallback: ").bold = True
            p.add_run(r_["fallback"])
            for s in (r_["flagged"] or r_["found"])[:2]:
                doc.add_paragraph(s[:600], style="Intense Quote")
    if terms:
        doc.add_heading("Key terms detected (keyword extraction — verify)", level=1)
        table = doc.add_table(rows=1, cols=3)
        table.style = "Light Grid Accent 1"
        for i, h in enumerate(["Type", "Value", "Context"]):
            table.rows[0].cells[i].text = h
        for t in terms[:40]:
            row = table.add_row().cells
            row[0].text, row[1].text, row[2].text = t["Type"], t["Value"], t["Context"][:220]
    doc.add_heading("Advocate review record", level=1)
    reviewer = m["lawyer_name"] or "Not recorded"
    if m["reviewer_username"]:
        reviewer += f" (account: {m['reviewer_username']}; enrolment no. as entered: {m['reviewer_enrolment'] or 'not entered'})"
    doc.add_paragraph(f"Reviewer: {reviewer}")
    doc.add_paragraph(f"Decision: {m['lawyer_decision']}")
    doc.add_paragraph(f"Notes: {m['lawyer_notes'] or 'No notes recorded.'}")
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def make_draft_docx(m, d):
    doc = Document()
    doc.add_heading(f"{d['kind']} — {m['matter_title']}", 0)
    reviewed = d["status"] == DRAFT_REVIEWED
    r = doc.add_paragraph().add_run(
        f"Reviewed by {d['reviewed_by']} on {d['reviewed_at']}. Not legal advice until the advocate adopts it."
        if reviewed else
        "UNREVIEWED AI DRAFT — NOT LEGAL ADVICE — NOT FOR CLIENT DELIVERY"
    )
    r.bold = True
    r.font.color.rgb = RGBColor(36, 117, 75) if reviewed else RGBColor(170, 72, 55)
    for line in (d["content"] or "").splitlines():
        if not line.strip():
            continue
        if line.startswith("#"):
            doc.add_heading(line.lstrip("#").strip().replace("**", ""), level=2)
        elif line.lstrip().startswith(("- ", "* ")):
            doc.add_paragraph(line.lstrip()[2:].replace("**", ""), style="List Bullet")
        else:
            doc.add_paragraph(line.replace("**", ""))
    quotes = json.loads(d["quotes_json"] or "[]")
    if quotes:
        doc.add_heading("Source quotes", level=1)
        for q in quotes:
            doc.add_paragraph(("[matched] " if q["found"] else "[NOT MATCHED — verify] ") + q["quote"])
    gaps = json.loads(d["gaps_json"] or "[]")
    if gaps:
        doc.add_heading("Not located in the supplied text", level=1)
        for g in gaps:
            doc.add_paragraph(g, style="List Bullet")
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def export_json():
    with db() as conn:
        data = {t: [dict(r) for r in conn.execute(f"SELECT * FROM {t}")]
                for t in ["matters", "findings", "drafts", "playbook", "activity", "audit_log"]}
        data["users"] = [dict(r) for r in conn.execute(
            "SELECT id,username,full_name,role,enrolment_no,active,created_at,last_login FROM users")]
    data["exported_at"] = now()
    return json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8")


# ================================================================ UI HELPERS
def go_to(page_name):
    st.session_state["page"] = page_name


def open_matter(mid):
    st.session_state["page"] = "Matter workspace"
    st.session_state["ws_matter_id"] = mid


def render_draft_body(content, quotes, gaps, truncated=False):
    if truncated:
        st.warning(f"The contract text exceeded {MAX_CHARS:,} characters, so the AI only saw the first part.")
    st.markdown(content)
    if quotes:
        found = sum(1 for q in quotes if q["found"])
        st.markdown(f"**Source quotes — {found} of {len(quotes)} matched the contract text**")
        for q in quotes:
            st.markdown(("✅ " if q["found"] else "⚠️ not matched: ") + f"“{q['quote']}”")
        if found < len(quotes):
            st.warning("Some quotes could not be matched to the stored text. Check them against the source before relying on this draft.")
    if gaps:
        st.markdown("**Not located in the supplied text**")
        for g in gaps:
            st.markdown(f"- {g}")


# =================================================================== APP SHELL
render_header()
user = current_user()
if not user:
    render_login()
    st.stop()

pages = PAGES_BASE + (["Admin"] if user["role"] == "admin" else [])
if st.session_state.get("page") not in pages:
    st.session_state["page"] = pages[0]

st.sidebar.markdown("### LEXPILOT")
st.sidebar.caption(f"{user['full_name']} · {ROLE_LABEL[user['role']]}")
page = st.sidebar.radio("WORKSPACE", pages, key="page")
st.sidebar.divider()
if st.sidebar.button("Sign out", use_container_width=True):
    audit("Sign out")
    st.session_state.clear()
    st.rerun()
st.sidebar.caption(f"v{APP_VERSION} · SQLite · Rules + optional Gemini")
st.sidebar.caption("Not legal advice · Advocate review required")

if user.get("must_change_pw"):
    st.warning("Your password was set by an administrator. Please change it under **My account**.")

# ===================================================================== OVERVIEW
if page == "Overview":
    matters = list_matters()
    total = len(matters)
    pending = sum(1 for x in matters if x["lawyer_decision"] == "Pending")
    done = sum(1 for x in matters if x["lawyer_decision"] == APPROVED)
    with db() as conn:
        high_count = conn.execute("SELECT COUNT(*) FROM findings WHERE risk_level='High'").fetchone()[0]
        risk_rows = conn.execute("SELECT risk_level AS Risk, COUNT(*) AS Findings FROM findings GROUP BY risk_level").fetchall()

    st.markdown('<div class="section-kicker">Your command centre</div>', unsafe_allow_html=True)
    st.title("Good work starts with clarity.")
    st.caption("A single workspace for incoming contracts, preliminary issue spotting and professional review.")

    cols = st.columns(4)
    for c, (label, value, foot) in zip(cols, [
        ("Total matters", total, "All uploaded assignments"),
        ("Awaiting review", pending, "Require advocate attention"),
        ("High-priority flags", high_count, "Across all matters"),
        ("Review completed", done, "Recorded in workflow"),
    ]):
        c.markdown(
            f'<div class="metric-card"><div class="metric-label">{label}</div>'
            f'<div class="metric-value">{value}</div><div class="metric-foot">{foot}</div></div>',
            unsafe_allow_html=True,
        )
    st.write("")
    left, right = st.columns([1.5, 1])
    with left:
        st.markdown("### Recent matters")
        if matters:
            df = pd.DataFrame([{
                "Matter ID": f"LP-{m['id']:04d}", "Client": m["client_name"], "Matter": m["matter_title"],
                "Method": m["analysis_mode"], "Status": m["status"], "Date": m["created_at"][:16],
            } for m in matters[:12]])
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.markdown(
                '<div class="soft-card"><b>No matters yet</b>'
                '<p class="small-muted">Upload a contract to start your first review.</p></div>',
                unsafe_allow_html=True,
            )
            st.write("")
            st.button("Start first review", type="primary", on_click=go_to, args=("New contract review",))
    with right:
        st.markdown("### Findings by risk level")
        if risk_rows:
            rdf = pd.DataFrame([dict(r) for r in risk_rows]).set_index("Risk")
            st.bar_chart(rdf)
        else:
            st.caption("Charts appear once a contract has been analysed.")
        st.markdown("### Workflow")
        st.markdown(
            """
        <div class="soft-card">
          <p><span class="pill pill-gold">1 · Intake</span> <b>Upload the agreement</b></p>
          <p><span class="pill">2 · Analyse</span> <b>Rules, playbook and optional AI</b></p>
          <p><span class="pill">3 · Verify</span> <b>Advocate review</b></p>
          <p><span class="pill pill-green">4 · Prepare</span> <b>Export the review pack</b></p>
        </div>""",
            unsafe_allow_html=True,
        )
    st.markdown(
        '<div class="footer-note">LEXPILOT · Built for disciplined contract operations, not automated legal advice.</div>',
        unsafe_allow_html=True,
    )

# ================================================================ NEW REVIEW
elif page == "New contract review":
    st.markdown('<div class="section-kicker">Matter intake</div>', unsafe_allow_html=True)
    st.title("Start a contract review")
    st.caption("Add the assignment details, upload a document and choose the preliminary analysis method.")

    ai_ok, ai_msg = ai_gate()
    engines = ["Rules-based preliminary triage"] + (["Gemini AI + rules safety-net"] if ai_ok else [])
    if not ai_ok:
        st.info(f"AI analysis is unavailable: {ai_msg}")
    ocr_ok = ocr_ready()

    with st.form("intake_form", clear_on_submit=False):
        c1, c2 = st.columns(2)
        client = c1.text_input("Client / business name", placeholder="Example: Northstar Technologies Pvt Ltd")
        matter_title = c2.text_input("Matter title", placeholder="Example: SaaS vendor agreement")
        c3, c4 = st.columns(2)
        ctype = c3.selectbox("Contract category", CONTRACT_TYPES)
        jurisdiction = c4.selectbox("Primary jurisdiction", ["India", "Karnataka, India", "Other / to be confirmed"])
        uploaded = st.file_uploader(
            f"Upload contract document (max {MAX_UPLOAD_MB} MB)", type=["pdf", "docx", "txt"],
            help="PDF, Word (.docx) or plain text. Scanned PDFs need OCR to be installed.",
        )
        mode = st.radio("Analysis engine", engines, horizontal=True)
        o1, o2 = st.columns(2)
        use_ocr = o1.checkbox("Run OCR on scanned PDF pages", value=ocr_ok, disabled=not ocr_ok,
                              help="Requires Tesseract OCR on the machine." if not ocr_ok else "Pages with no selectable text are read with OCR (slower).")
        redact = o2.checkbox("Mask obvious identifiers before sending to AI", value=get_setting("redact_default", "1") == "1",
                             help="Masks emails, phone numbers, PAN, GSTIN and Aadhaar-style numbers. Names and addresses are not masked.")
        consent = st.checkbox("I am authorised to process this document in this system.")
        ai_consent = st.checkbox("If I chose Gemini: I am permitted to send this document to the external AI provider.")
        submitted = st.form_submit_button("Analyse contract", type="primary", use_container_width=True)

    if submitted:
        use_ai = mode.startswith("Gemini")
        if not client.strip() or not matter_title.strip() or not uploaded:
            st.error("Enter the client, matter title and upload a document.")
        elif not consent:
            st.error("Confirm that you are authorised to process this document.")
        elif use_ai and not ai_consent:
            st.error("To use AI analysis, confirm you are permitted to send this document to the external AI provider.")
        elif uploaded.size > MAX_UPLOAD_MB * 1024 * 1024:
            st.error(f"The file is {uploaded.size / 1048576:.1f} MB. The limit is {MAX_UPLOAD_MB} MB.")
        else:
            raw = uploaded.getvalue()
            digest = hashlib.sha256(raw).hexdigest()
            name = uploaded.name.lower()
            try:
                info = {}
                with st.spinner("Reading the document..."):
                    if name.endswith(".pdf"):
                        extracted, info = extract_pdf(raw, use_ocr=use_ocr and ocr_ok)
                    elif name.endswith(".docx"):
                        extracted = extract_docx(raw)
                    else:
                        extracted = extract_txt(raw)
                extracted = clean_text(extracted)
                if len(extracted) < 80:
                    st.error("Not enough readable text was found. If this is a scanned PDF, install Tesseract OCR and tick the OCR option.")
                else:
                    if info.get("blank_pages"):
                        st.warning(f"{info['blank_pages']} page(s) had little or no text and may be scanned images — review the source manually.")
                    if info.get("ocr_pages"):
                        st.info(f"OCR was used on {info['ocr_pages']} page(s). OCR can misread text; verify against the original.")
                    dup = find_duplicate(digest)
                    if dup:
                        st.warning(f"This exact file was already uploaded as LP-{dup['id']:04d} ({dup['matter_title']}). A new matter was still created.")
                    findings = rule_analysis(extracted)
                    analysis_mode, ai_summary = "Rules", None
                    if use_ai:
                        with st.spinner("Running preliminary AI issue spotting..."):
                            ai_data, error = gemini_analysis(extracted, redact=redact)
                        if ai_data and ai_data["findings"]:
                            findings = combine_findings(ai_data["findings"], findings)
                            analysis_mode = "Gemini AI + rules"
                            ai_summary = ai_data.get("summary")
                            if ai_data.get("truncated"):
                                st.warning("The contract was longer than the AI limit; only the first part was analysed.")
                        else:
                            st.warning((error or "AI returned no findings.") + " Saved the rules-based result instead.")
                    mid = create_matter({
                        "client": client.strip(), "title": matter_title.strip(), "ctype": ctype,
                        "jurisdiction": jurisdiction, "filename": uploaded.name, "digest": digest,
                        "mode": analysis_mode, "ocr_used": 1 if info.get("ocr_pages") else 0,
                        "redacted": 1 if (use_ai and redact) else 0, "summary": ai_summary,
                    }, extracted, findings)
                    st.success(f"Matter LP-{mid:04d} created with {len(findings)} preliminary finding(s).")
                    if ai_summary:
                        st.info("AI summary (unverified): " + ai_summary)
                    st.button("Open matter workspace", on_click=open_matter, args=(mid,))
            except ValueError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Document processing failed: {type(exc).__name__}: {exc}. The file may be damaged or in an unsupported format.")

# ============================================================ MATTER WORKSPACE
elif page == "Matter workspace":
    st.markdown('<div class="section-kicker">Review and quality control</div>', unsafe_allow_html=True)
    st.title("Matter workspace")
    matters = list_matters()
    if not matters:
        st.info("No matters yet. Start a contract review to create one.")
        st.button("Start a review", type="primary", on_click=go_to, args=("New contract review",))
    else:
        search = st.text_input("Search matters", placeholder="Client, matter title or file name")
        shown_matters = [m for m in matters if not search.strip() or search.strip().lower() in
                         f"{m['client_name']} {m['matter_title']} {m['filename']}".lower()]
        if not shown_matters:
            st.warning("No matters match your search.")
            st.stop()
        labels = {m["id"]: f"LP-{m['id']:04d} · {m['client_name']} · {m['matter_title']}" for m in shown_matters}
        if st.session_state.get("ws_matter_id") not in labels:
            st.session_state.pop("ws_matter_id", None)
        mid = st.selectbox("Choose matter", list(labels), format_func=lambda i: labels[i], key="ws_matter_id")
        m, findings, events = load_matter(mid)

        st.markdown(f"### {m['matter_title']}")
        st.caption(f"{m['client_name']} · {m['contract_type']} · {m['filename']} · {m['created_at'][:16]}")
        a, b, c = st.columns(3)
        a.metric("Findings", len(findings))
        b.metric("Analysis", m["analysis_mode"])
        c.metric("Professional review", m["lawyer_decision"])
        if m["ocr_used"]:
            st.warning("Some pages were read with OCR. OCR can misread text, so verify quoted wording against the original document.")

        tab_f, tab_p, tab_k, tab_s, tab_a, tab_r, tab_l = st.tabs(
            ["Findings", "Playbook check", "Key terms", "Source document", "Assistant", "Advocate review", "Activity"])

        # ---- Findings
        with tab_f:
            if m["ai_summary"]:
                st.info("AI summary (unverified): " + m["ai_summary"])
            risk_filter = st.multiselect("Filter by risk level", RISK_LEVELS, default=RISK_LEVELS, key=f"rf_{mid}")
            shown = [f for f in findings if f["risk_level"] in risk_filter]
            if not shown:
                st.info("No findings match the filter.")
            for i, f in enumerate(shown, 1):
                risk = f["risk_level"]
                icon = "🔴" if risk == "High" else ("🟡" if risk == "Review" else "⚪")
                with st.expander(f"{icon} {i:02d} · {f['clause_name']} · {risk}", expanded=(i == 1)):
                    st.markdown("**Why this may matter**")
                    st.write(f["reason"])
                    st.markdown("**Suggested advocate review action**")
                    st.write(f["suggested_action"])
                    st.markdown("**Source location**")
                    st.caption(f["source_location"])
                    st.markdown("**Confidence / basis**")
                    st.caption(f["confidence"])
                    st.markdown("**Matched wording**")
                    if f["source_excerpt"]:
                        st.code(f["source_excerpt"], language=None)
                    else:
                        st.info("No source excerpt captured. Review the entire source document.")

        pb_items = list_playbook()
        pb_results = run_playbook(m["extracted_text"] or "", m["contract_type"], pb_items)
        terms = key_terms(m["extracted_text"] or "")

        # ---- Playbook
        with tab_p:
            st.caption(f"Checked against the active playbook for “{m['contract_type']}”. Edit positions on the Playbooks page.")
            if not pb_results:
                st.info("No active playbook items apply to this contract type.")
            else:
                if any(not r["approved"] for r in pb_results):
                    st.warning("Positions marked “starter” have not been approved by an advocate. Treat them as discussion prompts only.")
                st.dataframe(pd.DataFrame([{
                    "Flag": r["icon"], "Clause": r["clause"], "Result": r["status"],
                    "Position": "Approved" if r["approved"] else "Starter",
                } for r in pb_results]), use_container_width=True, hide_index=True)
                for r in pb_results:
                    with st.expander(f"{r['icon']} {r['clause']} — {r['status']}"):
                        if r["flagged"]:
                            st.markdown("**Red-flag wording**")
                            for s in r["flagged"]:
                                st.code(s, language=None)
                        if r["found"]:
                            st.markdown("**Wording located**")
                            for s in r["found"]:
                                st.code(s, language=None)
                        if not r["found"] and not r["flagged"]:
                            st.caption("Nothing matched the detection patterns. This does not prove the clause is absent — check the full document.")
                        tag = f"approved by {r['approved_by']}" if r["approved"] else "starter, not approved"
                        st.markdown(f"**Preferred position** ({tag})")
                        st.write(r["preferred"] or "—")
                        st.markdown("**Fallback**")
                        st.write(r["fallback"] or "—")

        # ---- Key terms
        with tab_k:
            st.caption("Detected by pattern matching (periods, dates, amounts, percentages). Verify each against the contract.")
            if terms:
                st.dataframe(pd.DataFrame(terms), use_container_width=True, hide_index=True)
            else:
                st.info("No time periods, dates, amounts or percentages were detected.")

        # ---- Source
        with tab_s:
            st.caption("Page and paragraph markers are retained where extraction supports them.")
            st.text_area("Document text", m["extracted_text"], height=480, disabled=True, label_visibility="collapsed")

        # ---- Assistant
        with tab_a:
            st.warning(
                "AI drafts are preliminary working material. They can be wrong or incomplete, and nothing here is "
                "legal advice until a qualified advocate has checked it against the full contract."
            )
            gate_ok, gate_msg = ai_gate()
            if not gate_ok:
                st.info(gate_msg + " Previously saved drafts stay available below.")
            if len(m["extracted_text"] or "") > MAX_CHARS:
                st.warning(f"This contract is longer than {MAX_CHARS:,} characters; the AI will only see the first part.")
            kind = st.selectbox("What should the assistant prepare?", list(DRAFT_KINDS), key=f"kind_{mid}")
            question = ""
            if kind == "Ask a question":
                question = st.text_area("Your question about this contract", height=90, key=f"q_{mid}",
                                        placeholder="Example: Who can terminate, and on how much notice?")
            redact_ai = st.checkbox("Mask obvious identifiers before sending", value=get_setting("redact_default", "1") == "1",
                                    key=f"red_{mid}", help="Masks emails, phone numbers, PAN, GSTIN and Aadhaar-style numbers. Names and addresses are not masked.")
            ai_consent = st.checkbox("I am permitted to send this contract to the external AI provider", key=f"cons_{mid}")
            if st.button("Generate draft", type="primary", disabled=not (gate_ok and ai_consent), key=f"gen_{mid}"):
                if kind == "Ask a question" and not question.strip():
                    st.error("Enter a question first.")
                else:
                    with st.spinner("Preparing draft from the contract text..."):
                        res, err = gemini_assistant(m["extracted_text"], kind, question, redact=redact_ai, matter_ref=mid)
                    if res:
                        st.session_state[f"preview_{mid}"] = (kind, question, res)
                    else:
                        st.error(err)

            pending_draft = st.session_state.get(f"preview_{mid}")
            if pending_draft:
                pkind, pquestion, pres = pending_draft
                st.markdown(f"#### Preview — {pkind}")
                with st.container(border=True):
                    render_draft_body(pres["content"], pres["quotes"], pres["gaps"], pres["truncated"])
                p1, p2 = st.columns(2)
                if p1.button("Save to matter", type="primary", key=f"savedraft_{mid}", use_container_width=True):
                    save_draft(mid, pkind, pquestion, pres)
                    st.session_state.pop(f"preview_{mid}", None)
                    st.toast("Draft saved to this matter.", icon="✅")
                    st.rerun()
                if p2.button("Discard", key=f"discard_{mid}", use_container_width=True):
                    st.session_state.pop(f"preview_{mid}", None)
                    st.rerun()

            st.markdown("#### Saved drafts")
            drafts = list_drafts(mid)
            if not drafts:
                st.caption("No drafts saved for this matter yet.")
            for d in drafts:
                reviewed = d["status"] == DRAFT_REVIEWED
                icon = "🟢" if reviewed else "🟠"
                with st.expander(f"{icon} {d['kind']} · {d['created_at'][:16]} · {d['status']}"):
                    if d["question"]:
                        st.caption(f"Question: {d['question']}")
                    render_draft_body(d["content"], json.loads(d["quotes_json"] or "[]"),
                                      json.loads(d["gaps_json"] or "[]"), bool(d["truncated"]))
                    st.divider()
                    if reviewed:
                        st.success(f"Marked reviewed by {d['reviewed_by']} on {d['reviewed_at'][:16]}.")
                    elif can("review"):
                        ck = st.checkbox("I have checked this draft against the source contract", key=f"rvck_{d['id']}")
                        if st.button("Mark as reviewed", key=f"rvbtn_{d['id']}", disabled=not ck):
                            review_draft(d["id"], mid, user)
                            st.rerun()
                    else:
                        st.caption("Only an advocate or administrator can mark a draft as reviewed.")
                    x1, x2 = st.columns(2)
                    x1.download_button("Download draft (.docx)", make_draft_docx(m, d),
                                       file_name=f"LexPilot_LP-{mid:04d}_draft_{d['id']}.docx",
                                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                       key=f"dl_{d['id']}", use_container_width=True)
                    if x2.button("Delete draft", key=f"deld_{d['id']}", use_container_width=True):
                        delete_draft(d["id"], mid)
                        st.rerun()

        # ---- Advocate review
        with tab_r:
            st.warning(
                "Review decisions are tied to the signed-in account. The enrolment number is typed in by the user and "
                "is not verified by this system."
            )
            if can("review"):
                with st.form(f"review_{mid}"):
                    st.markdown(f"**Reviewer:** {user['full_name']} ({user['username']})"
                                + (f" · Enrolment no. {user['enrolment_no']}" if user.get("enrolment_no") else ""))
                    notes = st.text_area("Review notes, corrections and instructions", value=m["lawyer_notes"] or "", height=180)
                    current = m["lawyer_decision"] if m["lawyer_decision"] in REVIEW_OPTIONS else "Pending"
                    decision = st.selectbox("Review status", REVIEW_OPTIONS, index=REVIEW_OPTIONS.index(current))
                    saved = st.form_submit_button("Save review record", type="primary")
                if saved:
                    save_review(mid, user, notes, decision)
                    st.toast("Review record saved.", icon="✅")
                    st.rerun()
            else:
                st.info("Only an advocate or administrator can record review decisions. The current record is shown below.")
                st.write(f"**Status:** {m['lawyer_decision']}")
                st.write(f"**Reviewer:** {m['lawyer_name'] or 'Not recorded'}")
                st.write(f"**Notes:** {m['lawyer_notes'] or 'No notes recorded.'}")
            if can("delete_matter"):
                with st.expander("Danger zone"):
                    st.caption("Permanently deletes this matter, its findings, drafts and activity log. The deletion itself is recorded in the audit log.")
                    confirm_del = st.checkbox("I understand this cannot be undone", key=f"del_confirm_{mid}")
                    if st.button("Delete matter", disabled=not confirm_del, key=f"del_{mid}"):
                        delete_matter(mid)
                        st.toast(f"Matter LP-{mid:04d} deleted.", icon="🗑️")
                        st.rerun()

        # ---- Activity
        with tab_l:
            if events:
                st.dataframe(pd.DataFrame([{
                    "Time": e["created_at"][:19], "User": e["username"] or "—", "Event": e["event"], "Detail": e["detail"],
                } for e in events]), use_container_width=True, hide_index=True)
            else:
                st.caption("No activity events.")

        st.divider()
        d1, d2 = st.columns(2)
        d1.download_button(
            "Download review pack (.docx)", make_docx_pack(m, findings, pb_results, terms, user["full_name"]),
            file_name=f"LexPilot_LP-{mid:04d}_review_pack.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        d2.download_button(
            "Download report (.md)", report_markdown(m, findings, pb_results, terms, user["full_name"]),
            file_name=f"LexPilot_LP-{mid:04d}_report.md", mime="text/markdown", use_container_width=True,
        )

# ================================================================== PLAYBOOKS
elif page == "Playbooks":
    st.markdown('<div class="section-kicker">Advocate-approved positions</div>', unsafe_allow_html=True)
    st.title("Clause playbooks")
    st.caption("Positions the firm prefers for each clause. Every contract is compared against the playbook for its type.")
    editable = can("playbook_edit")
    if not editable:
        st.info("You can view the playbook. Only advocates and administrators can edit or approve positions.")
    st.warning("Positions marked “starter” were supplied with the app as discussion prompts. An advocate must review and approve them before they are relied on.")

    items = list_playbook()
    type_filter = st.selectbox("Show items for", ["All types"] + CONTRACT_TYPES)
    shown = [i for i in items if type_filter == "All types" or applicable(i, type_filter)]
    st.caption(f"{len(shown)} item(s)")

    for it in shown:
        status = f"approved by {it['approved_by']}" if it["approved_by"] else "starter / unapproved"
        flag = "🟢" if it["approved_by"] else "🟠"
        off = "" if it["active"] else " · inactive"
        with st.expander(f"{flag} {it['clause_name']} · {status}{off}"):
            if editable:
                with st.form(f"pb_{it['id']}"):
                    name = st.text_input("Clause name", value=it["clause_name"])
                    applies = st.multiselect("Applies to", ["All"] + CONTRACT_TYPES,
                                             default=[t for t in (it["applies_to"] or "All").split(";") if t])
                    detect = st.text_area("Detection patterns (one regular expression per line)", value=it["detect_patterns"] or "", height=90)
                    flags_ = st.text_area("Red-flag patterns (one per line)", value=it["flag_patterns"] or "", height=70)
                    cc1, cc2, cc3 = st.columns(3)
                    required = cc1.checkbox("Clause is expected", value=bool(it["required"]))
                    risk_m = cc2.selectbox("Risk if not located", ["Review", "High"], index=0 if it["risk_if_missing"] != "High" else 1)
                    active = cc3.checkbox("Active", value=bool(it["active"]))
                    pref = st.text_area("Preferred position", value=it["preferred_position"] or "", height=90)
                    fall = st.text_area("Fallback position", value=it["fallback_position"] or "", height=70)
                    approve = st.checkbox("I approve this position as the firm's standard", value=bool(it["approved_by"]))
                    sv, dl = st.columns(2)
                    do_save = sv.form_submit_button("Save item", type="primary", use_container_width=True)
                    do_del = dl.form_submit_button("Delete item", use_container_width=True)
                if do_save:
                    try:
                        save_playbook_item(it["id"], {
                            "clause_name": name, "applies_to": applies, "detect_patterns": detect,
                            "flag_patterns": flags_, "required": required, "risk_if_missing": risk_m,
                            "active": active, "preferred_position": pref, "fallback_position": fall,
                        }, approve, user)
                        st.toast("Playbook item saved.", icon="✅")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))
                if do_del:
                    delete_playbook_item(it["id"])
                    st.rerun()
            else:
                st.write(f"**Applies to:** {(it['applies_to'] or 'All').replace(';', ', ')}")
                st.write(f"**Preferred position:** {it['preferred_position'] or '—'}")
                st.write(f"**Fallback:** {it['fallback_position'] or '—'}")

    if editable:
        st.markdown("### Add a clause")
        with st.form("pb_new", clear_on_submit=True):
            name = st.text_input("Clause name")
            applies = st.multiselect("Applies to", ["All"] + CONTRACT_TYPES, default=["All"])
            detect = st.text_area("Detection patterns (one regular expression per line)", height=80,
                                  placeholder="non.?solicit\nsolicitation of (employees|customers)")
            flags_ = st.text_area("Red-flag patterns (optional)", height=60)
            n1, n2 = st.columns(2)
            required = n1.checkbox("Clause is expected")
            risk_m = n2.selectbox("Risk if not located", ["Review", "High"])
            pref = st.text_area("Preferred position", height=80)
            fall = st.text_area("Fallback position", height=60)
            approve = st.checkbox("I approve this position as the firm's standard")
            go = st.form_submit_button("Add to playbook", type="primary")
        if go:
            try:
                save_playbook_item(None, {
                    "clause_name": name, "applies_to": applies, "detect_patterns": detect, "flag_patterns": flags_,
                    "required": required, "risk_if_missing": risk_m, "active": True,
                    "preferred_position": pref, "fallback_position": fall,
                }, approve, user)
                st.toast("Clause added.", icon="✅")
                st.rerun()
            except ValueError as e:
                st.error(str(e))

# ============================================================ COMPARE VERSIONS
elif page == "Compare versions":
    st.markdown('<div class="section-kicker">Version control</div>', unsafe_allow_html=True)
    st.title("Compare two contract versions")
    st.caption("Pick two matters (for example the draft and the counterparty's revised version) to see what changed.")
    matters = list_matters()
    if len(matters) < 2:
        st.info("You need at least two matters to compare. Upload both versions as separate reviews.")
    else:
        labels = {m["id"]: f"LP-{m['id']:04d} · {m['client_name']} · {m['matter_title']} · {m['filename']}" for m in matters}
        ids = list(labels)
        c1, c2 = st.columns(2)
        a_id = c1.selectbox("Version A (earlier)", ids, format_func=lambda i: labels[i], index=min(1, len(ids) - 1))
        b_id = c2.selectbox("Version B (later)", ids, format_func=lambda i: labels[i], index=0)
        if a_id == b_id:
            st.warning("Choose two different matters.")
        else:
            ma, fa, _ = load_matter(a_id)
            mb, fb, _ = load_matter(b_id)
            diff = diff_texts(ma["extracted_text"], mb["extracted_text"])
            added = sum(1 for d in diff if d.startswith("+") and not d.startswith("+++"))
            removed = sum(1 for d in diff if d.startswith("-") and not d.startswith("---"))
            m1, m2, m3 = st.columns(3)
            m1.metric("Sentences added", added)
            m2.metric("Sentences removed", removed)
            m3.metric("Identical", "Yes" if not diff else "No")
            names_a = {f["clause_name"] for f in fa}
            names_b = {f["clause_name"] for f in fb}
            if names_a != names_b:
                st.markdown("**Change in flagged clauses**")
                if names_b - names_a:
                    st.write("Newly flagged in B: " + ", ".join(sorted(names_b - names_a)))
                if names_a - names_b:
                    st.write("No longer flagged in B: " + ", ".join(sorted(names_a - names_b)))
            if diff:
                st.markdown("**Changes** (lines starting with − were removed, + were added)")
                limited = diff[:3000]
                st.code("\n".join(limited), language="diff")
                if len(diff) > 3000:
                    st.caption("Showing the first 3,000 changed lines.")
                st.download_button("Download changes (.txt)", "\n".join(diff), file_name=f"LexPilot_LP-{a_id:04d}_vs_LP-{b_id:04d}.txt")
            else:
                st.success("The extracted text of both versions is identical.")
            st.caption("The comparison uses extracted text, so formatting changes are not shown and PDF line breaks can add noise.")

# ================================================================== MY ACCOUNT
elif page == "My account":
    st.markdown('<div class="section-kicker">Profile</div>', unsafe_allow_html=True)
    st.title("My account")
    st.write(f"**Name:** {user['full_name']}  \n**Username:** {user['username']}  \n**Role:** {ROLE_LABEL[user['role']]}")
    if user["role"] in ("advocate", "admin"):
        with st.form("profile_form"):
            enrol = st.text_input("Bar enrolment number (optional, not verified by this system)", value=user.get("enrolment_no") or "")
            sv = st.form_submit_button("Save profile")
        if sv:
            with db() as conn:
                conn.execute("UPDATE users SET enrolment_no=? WHERE id=?", (enrol.strip(), user["id"]))
            audit("Profile updated")
            st.toast("Profile saved.", icon="✅")
            st.rerun()
    st.markdown("### Change password")
    with st.form("pw_form", clear_on_submit=True):
        cur = st.text_input("Current password", type="password")
        n1 = st.text_input("New password", type="password", help=f"At least {MIN_PW_LEN} characters, letters and numbers.")
        n2 = st.text_input("Confirm new password", type="password")
        go = st.form_submit_button("Change password", type="primary")
    if go:
        row = get_user_by_id(user["id"])
        if not verify_pw(cur, row):
            st.error("Your current password is incorrect.")
        elif n1 != n2:
            st.error("The new passwords do not match.")
        else:
            try:
                set_password(user["id"], n1)
                audit("Password changed")
                st.success("Password changed.")
            except ValueError as e:
                st.error(str(e))

# ===================================================================== ADMIN
elif page == "Admin" and user["role"] == "admin":
    st.markdown('<div class="section-kicker">Administration</div>', unsafe_allow_html=True)
    st.title("Admin")
    t_users, t_ai, t_backup, t_audit = st.tabs(["Users", "AI & privacy", "Backup & export", "Audit log"])

    with t_users:
        users = list_users()
        st.dataframe(pd.DataFrame([{
            "Username": u["username"], "Name": u["full_name"], "Role": ROLE_LABEL[u["role"]],
            "Active": "Yes" if u["active"] else "No", "Last sign-in": (u["last_login"] or "—")[:16],
        } for u in users]), use_container_width=True, hide_index=True)

        st.markdown("#### Add a user")
        with st.form("add_user", clear_on_submit=True):
            a1, a2 = st.columns(2)
            un = a1.text_input("Username")
            fn = a2.text_input("Full name")
            a3, a4 = st.columns(2)
            role = a3.selectbox("Role", ROLES, format_func=lambda r: ROLE_LABEL[r])
            enrol = a4.text_input("Enrolment number (advocates, optional)")
            tmp = st.text_input("Temporary password", type="password", help=f"At least {MIN_PW_LEN} characters, letters and numbers. The user is asked to change it.")
            go = st.form_submit_button("Create user", type="primary")
        if go:
            try:
                create_user(un, fn, role, tmp, enrol, must_change=True)
                audit("User created", f"{un} ({role})")
                st.success("User created.")
                st.rerun()
            except ValueError as e:
                st.error(str(e))

        st.markdown("#### Manage a user")
        user_labels = {u["id"]: f"{u['username']} · {ROLE_LABEL[u['role']]}" for u in users}
        target_id = st.selectbox("User", list(user_labels), format_func=lambda i: user_labels[i])
        target = get_user_by_id(target_id)
        admins_active = sum(1 for u in users if u["role"] == "admin" and u["active"])
        is_last_admin = target["role"] == "admin" and target["active"] and admins_active <= 1
        m1, m2 = st.columns(2)
        with m1:
            with st.form("reset_pw", clear_on_submit=True):
                npw = st.text_input("New temporary password", type="password")
                rs = st.form_submit_button("Reset password")
            if rs:
                try:
                    set_password(target["id"], npw, must_change=True)
                    audit("Password reset", target["username"])
                    st.success("Password reset. The user must change it at next sign-in.")
                except ValueError as e:
                    st.error(str(e))
        with m2:
            new_role = st.selectbox("Role", ROLES, index=ROLES.index(target["role"]), format_func=lambda r: ROLE_LABEL[r], key=f"role_{target['id']}")
            if st.button("Update role", key=f"updrole_{target['id']}"):
                if is_last_admin and new_role != "admin":
                    st.error("You cannot remove the last active administrator.")
                else:
                    with db() as conn:
                        conn.execute("UPDATE users SET role=? WHERE id=?", (new_role, target["id"]))
                    audit("Role changed", f"{target['username']} → {new_role}")
                    st.rerun()
            label = "Deactivate user" if target["active"] else "Reactivate user"
            if st.button(label, key=f"toggle_{target['id']}"):
                if target["active"] and (is_last_admin or target["id"] == user["id"]):
                    st.error("You cannot deactivate yourself or the last active administrator.")
                else:
                    with db() as conn:
                        conn.execute("UPDATE users SET active=? WHERE id=?", (0 if target["active"] else 1, target["id"]))
                    audit(label, target["username"])
                    st.rerun()

    with t_ai:
        st.markdown("#### External AI (Gemini)")
        st.warning("When enabled, contract text is sent to Google's Gemini API. Confirm provider terms, client consent, confidentiality, privilege and data-residency requirements first. Prefer a Workspace/Enterprise or Cloud account over a free consumer key.")
        allow = st.checkbox("Allow external AI features", value=get_setting("allow_external_ai", "0") == "1")
        red_def = st.checkbox("Mask obvious identifiers by default", value=get_setting("redact_default", "1") == "1")
        st.caption(f"API key configured: {'yes' if get_api_key() else 'no'} · Per-user limit: {AI_HOURLY_LIMIT} AI requests per hour")
        if st.button("Save AI settings", type="primary"):
            set_setting("allow_external_ai", "1" if allow else "0")
            set_setting("redact_default", "1" if red_def else "0")
            audit("AI settings changed", f"allow={allow}; redact_default={red_def}")
            st.toast("Settings saved.", icon="✅")
            st.rerun()
        st.caption(f"OCR available on this machine: {'yes' if ocr_ready() else 'no (install Tesseract to read scanned PDFs)'}")

    with t_backup:
        st.warning("Backups and exports contain confidential contract text. Store them securely and delete copies you no longer need.")
        if st.button("Prepare backup files"):
            st.session_state["backup_db"] = backup_bytes()
            st.session_state["backup_json"] = export_json()
            audit("Backup prepared")
        if st.session_state.get("backup_db"):
            stamp = datetime.now().strftime("%Y%m%d_%H%M")
            b1, b2 = st.columns(2)
            b1.download_button("Download database backup (.db)", st.session_state["backup_db"],
                               file_name=f"lexpilot_backup_{stamp}.db", mime="application/octet-stream", use_container_width=True)
            b2.download_button("Download data export (.json)", st.session_state["backup_json"],
                               file_name=f"lexpilot_export_{stamp}.json", mime="application/json", use_container_width=True)
            st.caption("To restore, stop the app and replace lexpilot.db with the backup file.")

    with t_audit:
        with db() as conn:
            rows = conn.execute("SELECT ts,username,action,matter_id,detail FROM audit_log ORDER BY id DESC LIMIT 1000").fetchall()
        adf = pd.DataFrame([dict(r) for r in rows])
        if adf.empty:
            st.caption("No audit entries yet.")
        else:
            adf.columns = ["Time", "User", "Action", "Matter", "Detail"]
            st.dataframe(adf, use_container_width=True, hide_index=True)
            st.download_button("Download audit log (.csv)", adf.to_csv(index=False), file_name="lexpilot_audit_log.csv", mime="text/csv")

# ================================================================== ROADMAP
elif page == "Product roadmap":
    st.markdown('<div class="section-kicker">Product development</div>', unsafe_allow_html=True)
    st.title("From prototype to service platform")
    st.caption("What is built, and what must be validated before commercial use.")
    roadmap = [
        ("Contract intake + extraction", "Built", "PDF, DOCX and text, matter records, duplicate detection, OCR for scanned pages when installed."),
        ("Issue spotting", "Built", "Transparent rules engine plus optional Gemini analysis with excerpt verification and a rules safety-net."),
        ("Clause playbooks", "Built — basic", "Per-contract-type positions with red-flag detection. Needs advocate approval of every position."),
        ("Advocate review workflow", "Built", "Role-based review, decisions tied to accounts, drafts gated until reviewed."),
        ("Matter assistant", "Built", "Grounded Q&A and drafts with quote verification and review gating."),
        ("Key terms and version comparison", "Built — basic", "Periods, dates, amounts and sentence-level diff between versions."),
        ("Accounts, audit and backup", "Built — basic", "Hashed passwords, lockout, timeouts, audit log, backup and export."),
        ("Word tracked-changes redlines", "Next", "Proposed clause alternatives exported as true Word revisions."),
        ("Production security", "Required before real clients",
         "Multi-factor sign-in or SSO, encryption at rest, tenant isolation, managed database, tested backup and restore, retention policy, penetration test."),
        ("Professional and commercial readiness", "Required before launch",
         "Indian professional-regulation advice, engagement terms, privacy notices and client consent, pricing and delivery SLAs."),
    ]
    for i, (title, status, desc) in enumerate(roadmap, 1):
        pill = "pill-green" if status.startswith("Built") else "pill-gold"
        st.markdown(
            f"""
        <div class="soft-card" style="margin-bottom:10px;display:flex;gap:15px;align-items:flex-start">
          <div style="background:#f7eddd;color:#916b36;border-radius:12px;padding:10px 14px;font-weight:800">{i}</div>
          <div style="flex:1"><b>{title}</b><div class="small-muted">{desc}</div></div>
          <span class="pill {pill}">{status}</span>
        </div>""",
            unsafe_allow_html=True,
        )
    st.info("Do not upload real client documents until the production-security and professional-readiness items above are complete.")
