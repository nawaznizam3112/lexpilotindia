"""
LexPilot India — Contract Intelligence Platform (MVP 0.3, single file)

Run:
    pip install -r requirements.txt
    streamlit run app.py

Optional AI analysis: set GEMINI_API_KEY in .streamlit/secrets.toml or as an
environment variable. The app works fully without it (rules engine).
"""

import hashlib
import io
import json
import os
import re
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, RGBColor

APP_DIR = Path(__file__).parent
DB_PATH = APP_DIR / "lexpilot.db"

PAGES = ["Overview", "New contract review", "Matter workspace", "Product roadmap"]
APPROVED = "Reviewed — approved for delivery"
REVIEW_OPTIONS = ["Pending", "Review in progress", APPROVED]
RISK_LEVELS = ["High", "Review", "Informational", "Unassessed"]

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
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def now():
    return datetime.now().isoformat(timespec="seconds")


def init_db():
    with db() as conn:
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
        """
        )
        # Safe upgrades for databases created by earlier versions
        cols = {r["name"] for r in conn.execute("PRAGMA table_info(matters)")}
        for col, decl in {
            "jurisdiction": "TEXT DEFAULT 'India'",
            "analysis_mode": "TEXT DEFAULT 'Rules'",
            "lawyer_name": "TEXT",
            "lawyer_notes": "TEXT",
            "lawyer_decision": "TEXT DEFAULT 'Pending'",
            "reviewed_at": "TEXT",
        }.items():
            if col not in cols:
                conn.execute(f"ALTER TABLE matters ADD COLUMN {col} {decl}")
        fcols = {r["name"] for r in conn.execute("PRAGMA table_info(findings)")}
        for col, decl in {
            "source_location": "TEXT DEFAULT 'Extracted document text'",
            "confidence": "TEXT DEFAULT 'Rule match'",
        }.items():
            if col not in fcols:
                conn.execute(f"ALTER TABLE findings ADD COLUMN {col} {decl}")


init_db()


def log_event(mid, event, detail=""):
    with db() as conn:
        conn.execute(
            "INSERT INTO activity(matter_id,event,detail,created_at) VALUES(?,?,?,?)",
            (mid, event, detail, now()),
        )


# ======================================================== DOCUMENT PROCESSING
def extract_pdf(data: bytes) -> str:
    pages = []
    with fitz.open(stream=data, filetype="pdf") as doc:
        for i, page in enumerate(doc, start=1):
            txt = page.get_text("text")
            if txt.strip():
                pages.append(f"[PAGE {i}]\n{txt}")
    return "\n\n".join(pages)


def extract_docx(data: bytes) -> str:
    document = Document(io.BytesIO(data))
    parts = []
    for i, p in enumerate(document.paragraphs, start=1):
        if p.text.strip():
            parts.append(f"[PARAGRAPH {i}] {p.text}")
    for ti, table in enumerate(document.tables, start=1):
        for ri, row in enumerate(table.rows, start=1):
            parts.append(f"[TABLE {ti} ROW {ri}] " + " | ".join(c.text for c in row.cells))
    return "\n".join(parts)


def clean_text(text: str) -> str:
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# ==================================================================== RULES
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
    ("Automatic renewal", [r"automatic renewal", r"auto.?renew", r"renewal term"], "Review",
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
     [r"personal data", r"data protection", r"data processor", r"information security", r"security incident"],
     "Review",
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


# ======================================================== OPTIONAL GEMINI AI
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


def get_api_key():
    try:
        key = st.secrets.get("GEMINI_API_KEY", "")
        if key:
            return key
    except Exception:
        pass
    return os.getenv("GEMINI_API_KEY", "")


def gemini_analysis(text):
    api_key = get_api_key()
    if not api_key:
        return None, "Gemini API key not configured. Add GEMINI_API_KEY to Streamlit secrets or environment."
    try:
        from google import genai

        client = genai.Client(api_key=api_key)
        prompt = SYSTEM_PROMPT + "\n\nCONTRACT TEXT:\n" + text[:90000]
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
            config={"response_mime_type": "application/json", "temperature": 0.1},
        )
        raw = (response.text or "{}").strip()
        raw = re.sub(r"^```(?:json)?|```$", "", raw, flags=re.M).strip()
        data = json.loads(raw)
        findings = []
        for f in data.get("findings", [])[:15]:
            risk = f.get("risk_level", "Review")
            findings.append({
                "name": str(f.get("clause_name", "Unspecified issue"))[:160],
                "risk": risk if risk in RISK_LEVELS else "Review",
                "excerpt": str(f.get("source_excerpt", ""))[:2200],
                "reason": str(f.get("reason", ""))[:1200],
                "action": str(f.get("suggested_action", ""))[:1200],
                "location": str(f.get("source_location", "Not identified"))[:200],
                "confidence": str(f.get("confidence", "Not stated"))[:200],
            })
        return {"summary": str(data.get("summary", ""))[:3000], "findings": findings}, None
    except Exception as e:
        return None, f"AI analysis failed: {type(e).__name__}: {e}"


# ================================================================== STORAGE
def create_matter(client, title, ctype, jurisdiction, filename, digest, extracted, findings, mode):
    with db() as conn:
        cur = conn.execute(
            """INSERT INTO matters
            (client_name,matter_title,contract_type,jurisdiction,filename,file_hash,extracted_text,
             created_at,analysis_mode,status)
            VALUES(?,?,?,?,?,?,?,?,?,?)""",
            (client, title, ctype, jurisdiction, filename, digest, extracted, now(), mode,
             "Awaiting advocate review"),
        )
        mid = cur.lastrowid
        for f in findings:
            conn.execute(
                """INSERT INTO findings
                (matter_id,clause_name,risk_level,source_excerpt,reason,suggested_action,
                 source_location,confidence) VALUES(?,?,?,?,?,?,?,?)""",
                (mid, f["name"], f["risk"], f["excerpt"], f["reason"], f["action"],
                 f["location"], f["confidence"]),
            )
    log_event(mid, "Matter created", f"Analysis mode: {mode}; file: {filename}")
    return mid


def list_matters():
    with db() as conn:
        return conn.execute("SELECT * FROM matters ORDER BY id DESC").fetchall()


def find_duplicate(digest):
    with db() as conn:
        return conn.execute(
            "SELECT id, matter_title FROM matters WHERE file_hash=? ORDER BY id DESC LIMIT 1", (digest,)
        ).fetchone()


def load_matter(mid):
    with db() as conn:
        matter = conn.execute("SELECT * FROM matters WHERE id=?", (mid,)).fetchone()
        findings = conn.execute("SELECT * FROM findings WHERE matter_id=? ORDER BY id", (mid,)).fetchall()
        events = conn.execute("SELECT * FROM activity WHERE matter_id=? ORDER BY id DESC", (mid,)).fetchall()
    return matter, findings, events


def save_review(mid, lawyer, notes, decision):
    status = {
        APPROVED: "Advocate review complete",
        "Review in progress": "Advocate review in progress",
    }.get(decision, "Awaiting advocate review")
    with db() as conn:
        conn.execute(
            "UPDATE matters SET lawyer_name=?,lawyer_notes=?,lawyer_decision=?,reviewed_at=?,status=? WHERE id=?",
            (lawyer, notes, decision, now(), status, mid),
        )
    log_event(mid, "Advocate review updated", f"{decision}; reviewer field: {lawyer}")


def delete_matter(mid):
    with db() as conn:
        conn.execute("DELETE FROM matters WHERE id=?", (mid,))


# ================================================================== EXPORTS
def report_markdown(m, findings):
    out = [
        "# LexPilot — Preliminary Contract Review Pack", "",
        "> INTERNAL WORKING DOCUMENT. Not legal advice, not an advocate's opinion and not approved for "
        "client delivery unless the appropriate advocate has independently reviewed and authorised it.", "",
        f"**Client:** {m['client_name']}  ",
        f"**Matter:** {m['matter_title']}  ",
        f"**Contract type:** {m['contract_type']}  ",
        f"**Jurisdiction:** {m['jurisdiction']}  ",
        f"**File:** {m['filename']}  ",
        f"**Analysis method:** {m['analysis_mode']}  ",
        f"**Created:** {m['created_at']}", "",
        "## Limitations",
        "Automated analysis can be incomplete or wrong. A qualified advocate must examine the complete "
        "agreement, transaction facts, instructions and current applicable law. A missing match does not "
        "mean a clause is absent or a contract is safe.", "",
        "## Findings",
    ]
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
    out += [
        "## Advocate review record",
        f"**Reviewer:** {m['lawyer_name'] or 'Not recorded'}",
        f"**Decision:** {m['lawyer_decision']}",
        f"**Notes:** {m['lawyer_notes'] or 'No notes recorded.'}",
    ]
    return "\n".join(out)


def make_docx_pack(m, findings):
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
        ("Source file", m["filename"]), ("Analysis method", m["analysis_mode"]),
        ("Created", m["created_at"]),
    ]:
        doc.add_paragraph(f"{label}: {val or 'Not recorded'}")
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
    doc.add_heading("Advocate review record", level=1)
    doc.add_paragraph(f"Reviewer: {m['lawyer_name'] or 'Not recorded'}")
    doc.add_paragraph(f"Decision: {m['lawyer_decision']}")
    doc.add_paragraph(f"Notes: {m['lawyer_notes'] or 'No notes recorded.'}")
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ============================================================== NAVIGATION
def go_to(page_name):
    st.session_state["page"] = page_name


render_header()
st.sidebar.markdown("### LEXPILOT")
st.sidebar.caption("CONTRACT INTELLIGENCE · INDIA")
page = st.sidebar.radio("WORKSPACE", PAGES, key="page")
st.sidebar.divider()
st.sidebar.markdown("**MVP 0.3**")
st.sidebar.caption("SQLite · PDF/DOCX · Rules + optional Gemini")
st.sidebar.caption("Local prototype — no multi-user security")
st.sidebar.caption("Not legal advice · Advocate review required")

# ================================================================ OVERVIEW
if page == "Overview":
    matters = list_matters()
    total = len(matters)
    pending = sum(1 for x in matters if x["lawyer_decision"] == "Pending")
    done = sum(1 for x in matters if x["lawyer_decision"] == APPROVED)
    with db() as conn:
        high_count = conn.execute("SELECT COUNT(*) FROM findings WHERE risk_level='High'").fetchone()[0]

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
        st.markdown("### Workflow")
        st.markdown(
            """
        <div class="soft-card">
          <p><span class="pill pill-gold">1 · Intake</span><br><b>Upload the agreement</b><br><span class="small-muted">PDF or Word document</span></p>
          <p><span class="pill">2 · Analyse</span><br><b>Identify review points</b><br><span class="small-muted">Rules engine or optional AI</span></p>
          <p><span class="pill">3 · Verify</span><br><b>Advocate review</b><br><span class="small-muted">Human judgment and corrections</span></p>
          <p><span class="pill pill-green">4 · Prepare</span><br><b>Export review pack</b><br><span class="small-muted">Internal report for professional review</span></p>
        </div>""",
            unsafe_allow_html=True,
        )
    st.markdown(
        '<div class="footer-note">LEXPILOT · Built for disciplined contract operations, not automated legal advice.</div>',
        unsafe_allow_html=True,
    )

# ============================================================ NEW REVIEW
elif page == "New contract review":
    st.markdown('<div class="section-kicker">Matter intake</div>', unsafe_allow_html=True)
    st.title("Start a contract review")
    st.caption("Add the assignment details, upload a document and choose the preliminary analysis method.")

    with st.form("intake_form", clear_on_submit=False):
        c1, c2 = st.columns(2)
        client = c1.text_input("Client / business name", placeholder="Example: Northstar Technologies Pvt Ltd")
        matter_title = c2.text_input("Matter title", placeholder="Example: SaaS vendor agreement")
        c3, c4 = st.columns(2)
        ctype = c3.selectbox("Contract category", [
            "Vendor agreement", "Services agreement", "NDA", "SaaS / technology agreement",
            "Consulting agreement", "Employment agreement", "Distribution agreement", "Lease / licence", "Other",
        ])
        jurisdiction = c4.selectbox("Primary jurisdiction", ["India", "Karnataka, India", "Other / to be confirmed"])
        uploaded = st.file_uploader(
            "Upload contract document", type=["pdf", "docx"],
            help="Text-based PDF or DOCX. Scanned PDFs are not OCR-processed in this version.",
        )
        mode = st.radio("Analysis engine", ["Rules-based preliminary triage", "Gemini AI + rules fallback"],
                        horizontal=True)
        consent = st.checkbox(
            "I confirm this is a synthetic/sample document or that I am authorised to process it in this local prototype."
        )
        submitted = st.form_submit_button("Analyse contract", type="primary", use_container_width=True)

    if submitted:
        if not client.strip() or not matter_title.strip() or not uploaded:
            st.error("Enter the client, matter title and upload a document.")
        elif not consent:
            st.error("Confirm that you are authorised to process this document.")
        else:
            raw = uploaded.getvalue()
            digest = hashlib.sha256(raw).hexdigest()
            try:
                extracted = extract_pdf(raw) if uploaded.name.lower().endswith(".pdf") else extract_docx(raw)
                extracted = clean_text(extracted)
                if len(extracted) < 80:
                    st.error("Not enough selectable text was extracted. The PDF may be scanned; OCR is not included yet.")
                else:
                    dup = find_duplicate(digest)
                    if dup:
                        st.warning(f"This exact file was already uploaded as LP-{dup['id']:04d} ({dup['matter_title']}). A new matter was still created.")
                    findings = rule_analysis(extracted)
                    analysis_mode = "Rules"
                    ai_summary = None
                    if mode.startswith("Gemini"):
                        with st.spinner("Running preliminary AI issue spotting..."):
                            ai_data, error = gemini_analysis(extracted)
                        if ai_data and ai_data["findings"]:
                            findings = ai_data["findings"]
                            analysis_mode = "Gemini AI"
                            ai_summary = ai_data.get("summary")
                        else:
                            st.warning((error or "AI returned no findings.") + " Saved the rules-based result instead.")
                    mid = create_matter(client.strip(), matter_title.strip(), ctype, jurisdiction,
                                        uploaded.name, digest, extracted, findings, analysis_mode)
                    st.success(f"Matter LP-{mid:04d} created with {len(findings)} preliminary finding(s).")
                    if ai_summary:
                        st.info("AI summary: " + ai_summary)
                    st.button("Open matter workspace", on_click=go_to, args=("Matter workspace",))
            except Exception as exc:
                st.error(f"Document processing failed: {type(exc).__name__}: {exc}")

# ========================================================= MATTER WORKSPACE
elif page == "Matter workspace":
    st.markdown('<div class="section-kicker">Review and quality control</div>', unsafe_allow_html=True)
    st.title("Matter workspace")
    matters = list_matters()
    if not matters:
        st.info("No matters yet. Start a contract review to create one.")
        st.button("Start a review", type="primary", on_click=go_to, args=("New contract review",))
    else:
        labels = {f"LP-{m['id']:04d} · {m['client_name']} · {m['matter_title']}": m["id"] for m in matters}
        chosen = st.selectbox("Choose matter", list(labels.keys()))
        mid = labels[chosen]
        m, findings, events = load_matter(mid)

        st.markdown(f"### {m['matter_title']}")
        st.caption(f"{m['client_name']} · {m['contract_type']} · {m['filename']} · {m['created_at'][:16]}")
        a, b, c = st.columns(3)
        a.metric("Findings", len(findings))
        b.metric("Analysis", m["analysis_mode"])
        c.metric("Professional review", m["lawyer_decision"])

        tab1, tab2, tab3, tab4 = st.tabs(["Findings", "Source document", "Advocate review", "Activity"])

        with tab1:
            risk_filter = st.multiselect("Filter by risk level", RISK_LEVELS, default=RISK_LEVELS)
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

        with tab2:
            st.caption("Page and paragraph markers are retained where extraction supports them.")
            st.text_area("Document text", m["extracted_text"], height=480, disabled=True,
                         label_visibility="collapsed")

        with tab3:
            st.warning(
                "Recording a status here does not authenticate the reviewer or establish that an "
                "advocate-client relationship exists. Use only within a properly approved professional workflow."
            )
            with st.form(f"review_{mid}"):
                reviewer = st.text_input("Reviewing advocate / reviewer", value=m["lawyer_name"] or "")
                notes = st.text_area("Review notes, corrections and instructions",
                                     value=m["lawyer_notes"] or "", height=180)
                current = m["lawyer_decision"] if m["lawyer_decision"] in REVIEW_OPTIONS else "Pending"
                decision = st.selectbox("Review status", REVIEW_OPTIONS, index=REVIEW_OPTIONS.index(current))
                saved = st.form_submit_button("Save review record", type="primary")
            if saved:
                save_review(mid, reviewer, notes, decision)
                st.toast("Review record saved.", icon="✅")
                st.rerun()

            with st.expander("Danger zone"):
                st.caption("Permanently deletes this matter, its findings and its activity log.")
                confirm_del = st.checkbox("I understand this cannot be undone", key=f"del_confirm_{mid}")
                if st.button("Delete matter", disabled=not confirm_del, key=f"del_{mid}"):
                    delete_matter(mid)
                    st.toast(f"Matter LP-{mid:04d} deleted.", icon="🗑️")
                    st.rerun()

        with tab4:
            if events:
                st.dataframe(
                    pd.DataFrame([{"Time": e["created_at"][:19], "Event": e["event"], "Detail": e["detail"]}
                                  for e in events]),
                    use_container_width=True, hide_index=True,
                )
            else:
                st.caption("No activity events.")

        st.divider()
        d1, d2 = st.columns(2)
        d1.download_button(
            "Download review pack (.docx)", make_docx_pack(m, findings),
            file_name=f"LexPilot_LP-{mid:04d}_review_pack.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        d2.download_button(
            "Download report (.md)", report_markdown(m, findings),
            file_name=f"LexPilot_LP-{mid:04d}_report.md",
            mime="text/markdown", use_container_width=True,
        )

# ================================================================ ROADMAP
elif page == "Product roadmap":
    st.markdown('<div class="section-kicker">Product development</div>', unsafe_allow_html=True)
    st.title("From prototype to service platform")
    st.caption("The following stages should be built and validated before commercial use.")
    roadmap = [
        ("1", "Contract intake + extraction", "Built", "PDF/DOCX extraction, matter records and SQLite."),
        ("2", "Preliminary issue spotting", "Built", "Transparent rules engine and optional Gemini analysis."),
        ("3", "Advocate review workflow", "Built — basic", "Reviewer notes, decision status and activity history."),
        ("4", "Professional DOCX review pack", "Built — basic", "Export findings and review record to Word."),
        ("5", "Clause playbooks", "Next", "Advocate-approved positions by contract type and customer risk appetite."),
        ("6", "Redline generation", "Next", "Proposed clause alternatives with source traceability and Word comparison."),
        ("7", "OCR + page citations", "Next", "Scanned PDF support and more precise page-level evidence."),
        ("8", "Production security", "Required before real clients",
         "Authentication, tenant isolation, encryption, audit, backups and retention controls."),
        ("9", "Commercial operations", "Required before launch",
         "Approved engagement structure, intake controls, pricing and delivery SLAs."),
    ]
    for num, title, status, desc in roadmap:
        pill = "pill-green" if status.startswith("Built") else "pill-gold"
        st.markdown(
            f"""
        <div class="soft-card" style="margin-bottom:10px;display:flex;gap:15px;align-items:flex-start">
          <div style="background:#f7eddd;color:#916b36;border-radius:12px;padding:10px 14px;font-weight:800">{num}</div>
          <div style="flex:1"><b>{title}</b><div class="small-muted">{desc}</div></div>
          <span class="pill {pill}">{status}</span>
        </div>""",
            unsafe_allow_html=True,
        )
    st.info(
        "Gemini integration is optional. To enable it, add GEMINI_API_KEY to Streamlit secrets. "
        "Do not send confidential client contracts to an external AI provider until privacy, privilege, "
        "consent, contractual and security reviews are complete."
    )
