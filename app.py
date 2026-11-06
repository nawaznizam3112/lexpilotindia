import streamlit as st
import sqlite3
import hashlib
import re
import io
from datetime import datetime
from pathlib import Path

import pandas as pd
import fitz  # PyMuPDF
from docx import Document

APP_DIR = Path(__file__).parent
DB_PATH = APP_DIR / "lexpilot.db"

st.set_page_config(page_title="LexPilot | SME Contract Desk", page_icon="⚖️", layout="wide")

# -------------------- DATABASE --------------------
def db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with db() as conn:
        conn.executescript("""
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
            status TEXT DEFAULT 'AI analysis complete',
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
            FOREIGN KEY(matter_id) REFERENCES matters(id)
        );
        """)
init_db()

# -------------------- DOCUMENT EXTRACTION --------------------
def extract_pdf(data):
    doc = fitz.open(stream=data, filetype="pdf")
    return "\n".join(page.get_text("text") for page in doc)

def extract_docx(data):
    document = Document(io.BytesIO(data))
    chunks = [p.text for p in document.paragraphs if p.text.strip()]
    for table in document.tables:
        for row in table.rows:
            chunks.append(" | ".join(cell.text for cell in row.cells))
    return "\n".join(chunks)

def clean_text(text):
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

# -------------------- TRANSPARENT RULE-BASED TRIAGE --------------------
RULES = [
    {
        "name": "Limitation of liability",
        "patterns": [r"unlimited liability", r"without limitation", r"no limit.*liabilit", r"liabilit.*unlimited"],
        "risk": "High",
        "reason": "The wording may create uncapped or unusually broad liability exposure.",
        "action": "Ask an advocate to assess whether an appropriate liability cap, exclusions and carve-outs are needed."
    },
    {
        "name": "Indemnity",
        "patterns": [r"indemnif", r"hold harmless"],
        "risk": "Review",
        "reason": "Indemnity obligations can create significant financial exposure; scope, triggers and procedure matter.",
        "action": "Review covered losses, third-party claims, defence control, notice, mitigation and any liability cap interaction."
    },
    {
        "name": "Termination",
        "patterns": [r"terminat", r"terminate for convenience"],
        "risk": "Review",
        "reason": "Termination rights, notice periods and consequences may be asymmetric or incomplete.",
        "action": "Check mutuality, notice, cure periods, payment on termination, transition assistance and survival clauses."
    },
    {
        "name": "Automatic renewal",
        "patterns": [r"automatic renewal", r"auto.?renew", r"renewal term"],
        "risk": "Review",
        "reason": "An automatic renewal may create an unintended extension or notice deadline.",
        "action": "Confirm renewal duration, notice window, method of notice and a calendar reminder."
    },
    {
        "name": "Payment and late fees",
        "patterns": [r"payment terms", r"invoice", r"late fee", r"interest.*overdue", r"within \d+ days"],
        "risk": "Review",
        "reason": "Payment timing, invoice disputes, taxes and late-payment consequences affect cash flow.",
        "action": "Confirm due dates, invoice requirements, dispute process, taxes, withholding and late-payment calculation."
    },
    {
        "name": "Confidentiality",
        "patterns": [r"confidential", r"non.?disclosure"],
        "risk": "Review",
        "reason": "Confidentiality language should define protected information, permitted use and exceptions.",
        "action": "Check exclusions, permitted disclosures, duration, return/deletion and security obligations."
    },
    {
        "name": "Intellectual property",
        "patterns": [r"intellectual property", r"work product", r"ownership of deliverables", r"assign.*copyright"],
        "risk": "Review",
        "reason": "Ownership and licensing of pre-existing IP, deliverables and improvements may be unclear.",
        "action": "Separate background IP from project deliverables; confirm assignment/licence scope and payment conditions."
    },
    {
        "name": "Dispute resolution and governing law",
        "patterns": [r"arbitrat", r"governing law", r"jurisdiction", r"courts at"],
        "risk": "Review",
        "reason": "Dispute forum, governing law, seat/venue and escalation procedures should be consistent.",
        "action": "Ask an advocate to verify governing law, exclusive jurisdiction, arbitration seat, venue and appointment procedure."
    },
    {
        "name": "Data protection and security",
        "patterns": [r"personal data", r"data protection", r"data processor", r"information security", r"security incident"],
        "risk": "Review",
        "reason": "The contract may involve personal data or security duties that need specific allocation.",
        "action": "Review roles, permitted processing, security controls, incident notice, subprocessors, retention and deletion."
    },
    {
        "name": "Assignment and subcontracting",
        "patterns": [r"assign.*agreement", r"subcontract", r"sub.?contract"],
        "risk": "Review",
        "reason": "Transfer or subcontracting rights can change who performs or receives contractual obligations.",
        "action": "Check consent requirements, permitted group transfers and responsibility for subcontractors."
    },
    {
        "name": "Force majeure",
        "patterns": [r"force majeure", r"act of god"],
        "risk": "Review",
        "reason": "The clause may affect performance obligations during events outside a party's control.",
        "action": "Check event definition, notice, mitigation, suspension period and termination rights."
    },
]

def split_sentences(text):
    # Keep useful clause-sized excerpts without requiring external AI.
    parts = re.split(r"(?<=[.;])\s+|\n+", text)
    return [p.strip() for p in parts if len(p.strip()) > 35]

def analyse(text):
    findings = []
    sentences = split_sentences(text)
    lower = text.lower()
    for rule in RULES:
        hits = []
        for sentence in sentences:
            if any(re.search(p, sentence, re.I) for p in rule["patterns"]):
                hits.append(sentence[:1200])
        # For clause topics, absence is not automatically a legal defect.
        if hits:
            findings.append({
                "name": rule["name"],
                "risk": rule["risk"],
                "excerpt": "\n---\n".join(hits[:3]),
                "reason": rule["reason"],
                "action": rule["action"]
            })
    if not findings:
        findings.append({
            "name": "No rule-based flags detected",
            "risk": "Unassessed",
            "excerpt": "",
            "reason": "The prototype did not identify text matching its limited review rules. This does not mean the contract is safe or complete.",
            "action": "A qualified advocate should review the entire agreement and the transaction context."
        })
    return findings

def save_matter(client, title, ctype, filename, file_hash, extracted, findings):
    with db() as conn:
        cur = conn.execute(
            """INSERT INTO matters(client_name,matter_title,contract_type,filename,file_hash,
            extracted_text,created_at) VALUES(?,?,?,?,?,?,?)""",
            (client, title, ctype, filename, file_hash, extracted, datetime.now().isoformat(timespec="seconds"))
        )
        mid = cur.lastrowid
        for f in findings:
            conn.execute(
                """INSERT INTO findings(matter_id,clause_name,risk_level,source_excerpt,reason,suggested_action)
                VALUES(?,?,?,?,?,?)""",
                (mid, f["name"], f["risk"], f["excerpt"], f["reason"], f["action"])
            )
        return mid

def get_matters():
    with db() as conn:
        return conn.execute("SELECT * FROM matters ORDER BY id DESC").fetchall()

def get_matter(mid):
    with db() as conn:
        m = conn.execute("SELECT * FROM matters WHERE id=?", (mid,)).fetchone()
        fs = conn.execute("SELECT * FROM findings WHERE matter_id=? ORDER BY id", (mid,)).fetchall()
    return m, fs

def update_review(mid, lawyer, notes, decision):
    status = "Advocate review complete" if decision == "Reviewed - approved for delivery" else "Advocate review in progress"
    with db() as conn:
        conn.execute("""UPDATE matters SET lawyer_name=?, lawyer_notes=?, lawyer_decision=?,
                     reviewed_at=?, status=? WHERE id=?""",
                     (lawyer, notes, decision, datetime.now().isoformat(timespec="seconds"), status, mid))

def make_report(m, findings):
    lines = [
        "# LexPilot – Contract Triage Report",
        "",
        "**Prototype status:** AI-assisted preliminary triage; not legal advice or an advocate's opinion.",
        f"**Client:** {m['client_name']}",
        f"**Matter:** {m['matter_title']}",
        f"**Contract type:** {m['contract_type'] or 'Not specified'}",
        f"**Source file:** {m['filename'] or 'Not recorded'}",
        f"**Created:** {m['created_at']}",
        "",
        "## Important limitation",
        "This report is generated using transparent keyword/rule matching. It does not understand the full legal or commercial context, may miss important issues, and may flag clauses that are acceptable in context. No absence of flags indicates legal safety. An independent qualified advocate must review the complete agreement and approve any advice or final legal deliverable.",
        "",
        "## Findings",
    ]
    for i, f in enumerate(findings, 1):
        lines += [
            f"### {i}. {f['clause_name']} — {f['risk_level']}",
            f"**Why it was flagged:** {f['reason']}",
            f"**Suggested review action:** {f['suggested_action']}",
            f"**Matched source text:**\n\n> {(f['source_excerpt'] or 'No matching text captured.')[:2500]}",
            ""
        ]
    lines += [
        "## Advocate review",
        f"**Reviewer:** {m['lawyer_name'] or 'Pending'}",
        f"**Decision:** {m['lawyer_decision']}",
        f"**Notes:** {m['lawyer_notes'] or 'No review notes recorded.'}",
        "",
        "This prototype report is an internal working document until reviewed and approved through the appropriate professional engagement."
    ]
    return "\n".join(lines)

# -------------------- UI --------------------
st.markdown("""
<style>
.block-container {padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1450px;}
.hero {padding: 1.35rem 1.5rem; border-radius: 18px; background: linear-gradient(120deg,#101c32,#263d5c); color: white; margin-bottom: 1rem;}
.hero h1 {color:white; margin:0; font-size:2rem;}
.hero p {color:#d7e2ef; margin:.4rem 0 0 0;}
.small-muted {color:#6b7280; font-size:.85rem;}
</style>
<div class="hero">
<h1>⚖️ LexPilot</h1>
<p>SME Contract Desk · Document intake · Clause intelligence · Advocate review workflow</p>
</div>
""", unsafe_allow_html=True)

st.warning("Prototype only. This version uses local rule-based text matching, not a validated legal AI model. It does not provide legal advice. Do not upload real client, privileged, personal or commercially sensitive documents until security, access controls, retention, consent and professional arrangements have been reviewed.")

page = st.sidebar.radio("Workspace", ["Dashboard", "New contract review", "Matter workspace", "How this prototype works"])
st.sidebar.caption("LexPilot MVP 0.1 · SQLite")

if page == "Dashboard":
    matters = get_matters()
    total = len(matters)
    pending = sum(1 for m in matters if m["lawyer_decision"] == "Pending")
    reviewed = sum(1 for m in matters if m["lawyer_decision"] == "Reviewed - approved for delivery")
    c1,c2,c3 = st.columns(3)
    c1.metric("Matters created", total)
    c2.metric("Awaiting advocate review", pending)
    c3.metric("Review completed", reviewed)
    st.subheader("Recent matters")
    if matters:
        df = pd.DataFrame([{
            "ID": m["id"], "Client": m["client_name"], "Matter": m["matter_title"],
            "Type": m["contract_type"], "Status": m["status"], "Created": m["created_at"][:16]
        } for m in matters])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("No matters yet. Create your first test matter from **New contract review**.")
    st.subheader("MVP workflow")
    st.write("1. Create matter → 2. Extract contract text → 3. Run preliminary rule-based triage → 4. Advocate reviews findings → 5. Export internal report.")

elif page == "New contract review":
    st.subheader("Create a contract review matter")
    with st.form("new_matter"):
        a,b = st.columns(2)
        client = a.text_input("Client / business name", placeholder="Example: ABC Technologies Pvt Ltd")
        title = b.text_input("Matter title", placeholder="Example: Vendor services agreement")
        ctype = st.selectbox("Contract type", ["Vendor agreement", "Services agreement", "NDA", "Employment agreement", "Consulting agreement", "Distribution agreement", "Other"])
        uploaded = st.file_uploader("Upload contract", type=["pdf","docx"])
        submitted = st.form_submit_button("Extract and analyse contract", type="primary", use_container_width=True)
    if submitted:
        if not client.strip() or not title.strip() or not uploaded:
            st.error("Please enter the client, matter title and upload a PDF or DOCX file.")
        else:
            data = uploaded.getvalue()
            digest = hashlib.sha256(data).hexdigest()
            try:
                if uploaded.name.lower().endswith(".pdf"):
                    extracted = extract_pdf(data)
                else:
                    extracted = extract_docx(data)
                extracted = clean_text(extracted)
                if len(extracted) < 80:
                    st.error("Very little text was extracted. This may be a scanned PDF. OCR is not included in this version.")
                else:
                    findings = analyse(extracted)
                    mid = save_matter(client.strip(), title.strip(), ctype, uploaded.name, digest, extracted, findings)
                    st.success(f"Matter #{mid} created and preliminary triage completed.")
                    st.info("Open Matter workspace to inspect the extracted text, matched clauses and advocate review section.")
            except Exception as e:
                st.error(f"Could not process this file: {e}")

elif page == "Matter workspace":
    matters = get_matters()
    if not matters:
        st.info("Create a matter first.")
    else:
        choices = {f"#{m['id']} · {m['client_name']} · {m['matter_title']}": m["id"] for m in matters}
        selected = st.selectbox("Select matter", list(choices.keys()))
        mid = choices[selected]
        m, findings = get_matter(mid)
        st.markdown(f"### {m['matter_title']}")
        st.caption(f"Client: {m['client_name']} · File: {m['filename']} · Created: {m['created_at']}")
        x,y,z = st.columns(3)
        x.metric("Findings", len(findings))
        y.metric("Advocate decision", m["lawyer_decision"])
        z.metric("Workflow status", m["status"])
        st.subheader("Preliminary findings")
        for f in findings:
            with st.expander(f"{f['risk_level']} · {f['clause_name']}", expanded=True):
                st.write("**Why flagged**")
                st.write(f["reason"])
                st.write("**Suggested review action**")
                st.write(f["suggested_action"])
                st.write("**Matched text**")
                st.code(f["source_excerpt"] or "No excerpt captured.")
        with st.expander("Extracted contract text"):
            st.text_area("Extracted text (read-only working copy)", m["extracted_text"], height=300, disabled=True)
        st.divider()
        st.subheader("Advocate review workspace")
        st.caption("This section records a review workflow. It does not itself verify the identity, qualification or authority of the reviewer.")
        with st.form(f"review_{mid}"):
            lawyer = st.text_input("Reviewing advocate", value=m["lawyer_name"] or "")
            notes = st.text_area("Advocate review notes", value=m["lawyer_notes"] or "", height=140)
            decision = st.selectbox("Review status", ["Pending", "Review in progress", "Reviewed - approved for delivery"], index=["Pending","Review in progress","Reviewed - approved for delivery"].index(m["lawyer_decision"]) if m["lawyer_decision"] in ["Pending","Review in progress","Reviewed - approved for delivery"] else 0)
            save = st.form_submit_button("Save review record", type="primary")
        if save:
            update_review(mid, lawyer, notes, decision)
            st.success("Review record saved.")
            st.rerun()
        m, findings = get_matter(mid)
        report = make_report(m, findings)
        st.download_button("Download internal review report (.md)", report, file_name=f"lexpilot_matter_{mid}_report.md", mime="text/markdown", use_container_width=True)

elif page == "How this prototype works":
    st.subheader("What this first version does")
    st.markdown("""
- Stores matters, extracted text, findings and review records in a local SQLite database.
- Extracts selectable text from PDF and DOCX files.
- Uses transparent keyword rules to identify clauses that may deserve closer attention.
- Preserves matched excerpts so a reviewer can return to the source wording.
- Provides an advocate review record and downloadable internal report.

### What it does not do
- It does not provide legal advice or determine whether a contract is enforceable.
- It does not generate legally validated replacement clauses.
- It does not perform OCR on scanned PDFs.
- It does not authenticate advocates or provide secure multi-user access.
- It does not encrypt the SQLite database or uploaded source files (source files are processed in memory and are not saved as files by this app).
- It does not replace review of the full agreement and transaction context.

### Next development stages
1. Add page/paragraph-level source references and OCR.
2. Add an AI provider behind a controlled, privacy-reviewed interface.
3. Add clause playbooks approved by advocates.
4. Add DOCX redlining and report generation.
5. Add secure authentication, tenant isolation, encryption and audit logs.
6. Add production deployment and a compliant client engagement workflow.
""")
    st.info("The database file lexpilot.db is created in the same folder as app.py. Keep it private and back it up securely.")

st.sidebar.divider()
st.sidebar.caption("Internal prototype · Not for unsupervised legal use")
