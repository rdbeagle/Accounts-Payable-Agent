"""
app.py — Integrity Wall Systems PO Automation Dashboard
Run with: streamlit run app.py
"""

import os
import json
import subprocess
import pandas as pd
import streamlit as st
from pathlib import Path
from datetime import datetime, date
from dotenv import load_dotenv

load_dotenv()

STORAGE_ROOT = os.getenv("STORAGE_ROOT", "./po-automation/data")
TRACKING_CSV = os.path.join(STORAGE_ROOT, "po_tracking.csv")

st.set_page_config(
    page_title="IWS — PO Automation",
    page_icon="🏗️",
    layout="wide",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;1,9..40,300&display=swap');

[data-testid="stAppViewContainer"] {
    background-color: #0c1228;
    background-image:
        linear-gradient(160deg, rgba(180,30,30,0.06) 0%, transparent 40%),
        repeating-linear-gradient(-52deg,
            transparent, transparent 58px,
            rgba(180,30,30,0.055) 58px, rgba(180,30,30,0.055) 60px),
        repeating-linear-gradient(38deg,
            transparent, transparent 58px,
            rgba(255,255,255,0.025) 58px, rgba(255,255,255,0.025) 60px);
    background-attachment: fixed;
}
[data-testid="stHeader"] { background: transparent !important; }
[data-testid="stSidebar"] {
    background: rgba(6,10,26,0.98) !important;
    border-right: 1px solid rgba(180,30,30,0.25) !important;
}
[data-testid="stSidebar"] > div { padding-top: 2rem; }

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif !important;
    color: rgba(255,255,255,0.85) !important;
}

.iws-logo-mark {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 3.5rem;
    line-height: 1;
    letter-spacing: 8px;
    color: #fff;
    text-shadow: 0 0 60px rgba(220,40,40,0.6), 0 0 20px rgba(220,40,40,0.3);
}
.iws-logo-sub {
    font-size: 0.62rem;
    letter-spacing: 5px;
    text-transform: uppercase;
    color: rgba(255,255,255,0.4);
    margin-top: 2px;
    font-weight: 300;
}
.red-bar {
    height: 2px;
    background: linear-gradient(90deg, #B41E1E 0%, #E63232 50%, transparent 100%);
    border: none;
    margin: 14px 0;
}
.section-head {
    font-size: 0.6rem;
    letter-spacing: 4px;
    text-transform: uppercase;
    color: rgba(255,255,255,0.6);
    font-weight: 400;
    margin-bottom: 10px;
    padding-bottom: 6px;
    border-bottom: 1px solid rgba(255,255,255,0.06);
}
.page-title {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 2.8rem;
    letter-spacing: 5px;
    color: #ffffff;
    line-height: 1;
    margin: 0;
}
.page-sub {
    font-size: 0.7rem;
    letter-spacing: 4px;
    text-transform: uppercase;
    color: rgba(255,255,255,0.6);
    margin-top: 4px;
    font-weight: 300;
}

/* AI Brief card */
.ai-brief-card {
    background: rgba(180,30,30,0.08);
    border: 1px solid rgba(180,30,30,0.25);
    border-left: 3px solid #E63232;
    border-radius: 10px;
    padding: 18px 22px;
    margin-bottom: 20px;
}
.ai-brief-label {
    font-size: 0.58rem;
    letter-spacing: 4px;
    text-transform: uppercase;
    color: rgba(230,50,50,0.7);
    font-weight: 500;
    margin-bottom: 8px;
}
.ai-brief-text {
    font-size: 0.92rem;
    color: rgba(255,255,255,0.85);
    line-height: 1.6;
    font-weight: 300;
}

/* KPI cards */
.kpi-wrap {
    background: rgba(255,255,255,0.03);
    border: 1px solid rgba(255,255,255,0.07);
    border-radius: 10px;
    padding: 22px 16px 18px;
    text-align: center;
    transition: border-color 0.2s;
}
.kpi-wrap:hover { border-color: rgba(255,255,255,0.14); }
.kpi-n { font-family: 'Bebas Neue', sans-serif; font-size: 3rem; line-height: 1; color: #fff; }
.kpi-l { font-size: 0.6rem; letter-spacing: 2px; text-transform: uppercase; color: rgba(255,255,255,0.55); margin-top: 5px; }
.kpi-clean  .kpi-n { color: #4ade80; }
.kpi-flag   .kpi-n { color: #f87171; }
.kpi-dup    .kpi-n { color: #E63232; }
.kpi-mis    .kpi-n { color: #fb923c; }

/* Alert cards */
.alert-card {
    background: rgba(180,30,30,0.1);
    border: 1px solid rgba(180,30,30,0.3);
    border-left: 3px solid #E63232;
    border-radius: 8px;
    padding: 14px 18px;
    margin-bottom: 10px;
    font-size: 0.83rem;
}
.alert-card.dup  { border-left-color: #E63232; background: rgba(180,30,30,0.12); }
.alert-card.mis  { border-left-color: #fb923c; background: rgba(251,146,60,0.08); border-color: rgba(251,146,60,0.25); }
.alert-card.err  { border-left-color: rgba(255,255,255,0.2); background: rgba(255,255,255,0.03); border-color: rgba(255,255,255,0.08); }
.badge {
    display: inline-block;
    padding: 2px 9px;
    border-radius: 4px;
    font-size: 0.65rem;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    font-weight: 500;
    margin-right: 8px;
}
.badge-dup { background: rgba(180,30,30,0.3); color: #fca5a5; }
.badge-mis { background: rgba(251,146,60,0.2); color: #fdba74; }
.badge-err { background: rgba(255,255,255,0.07); color: rgba(255,255,255,0.45); }
.badge-cln { background: rgba(74,222,128,0.15); color: #4ade80; }
.alert-meta { font-size: 0.73rem; color: rgba(255,255,255,0.4); margin-top: 6px; }
.alert-flag-text { color: rgba(255,255,255,0.65); margin-top: 4px; }

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, #9B1A1A, #D42B2B) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'Bebas Neue', sans-serif !important;
    letter-spacing: 3px !important;
    font-size: 0.95rem !important;
    padding: 10px 20px !important;
    box-shadow: 0 2px 16px rgba(180,30,30,0.25) !important;
    transition: all 0.15s !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #B42020, #E63232) !important;
    box-shadow: 0 4px 24px rgba(220,50,50,0.45) !important;
    transform: translateY(-1px) !important;
}
[data-testid="stFileUploader"] {
    border: 1px dashed rgba(180,30,30,0.35) !important;
    border-radius: 8px !important;
    background: rgba(255,255,255,0.02) !important;
}
[data-testid="stDataFrame"] thead tr th {
    background: rgba(255,255,255,0.04) !important;
    font-size: 0.7rem !important;
    letter-spacing: 1.5px !important;
    text-transform: uppercase !important;
    color: rgba(255,255,255,0.4) !important;
    font-weight: 400 !important;
}
[data-testid="stDataFrame"] tbody tr:hover td {
    background: rgba(255,255,255,0.04) !important;
}
[data-testid="stSelectbox"] > div,
[data-testid="stTextInput"] > div > div {
    background: rgba(255,255,255,0.04) !important;
    border-color: rgba(255,255,255,0.1) !important;
    border-radius: 6px !important;
}
.stDownloadButton > button {
    background: rgba(255,255,255,0.05) !important;
    color: rgba(255,255,255,0.7) !important;
    border: 1px solid rgba(255,255,255,0.1) !important;
    font-family: 'Bebas Neue', sans-serif !important;
    letter-spacing: 2px !important;
}
.stDownloadButton > button:hover {
    background: rgba(255,255,255,0.09) !important;
    border-color: rgba(255,255,255,0.2) !important;
}
</style>
""", unsafe_allow_html=True)


# ── Helpers ────────────────────────────────────────────────────────────────────
def excel_serial_to_date(val) -> str:
    try:
        f = float(val)
        if f > 1000:
            from datetime import timedelta
            return (date(1899, 12, 30) + timedelta(days=int(f))).strftime("%m/%d/%Y")
        return str(val)
    except Exception:
        return str(val) if val and str(val) != "nan" else "—"


def load_tracking() -> pd.DataFrame:
    path = Path(TRACKING_CSV)
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_csv(path)
    if "flags" in df.columns:
        df["flags"] = df["flags"].apply(
            lambda x: json.loads(x) if isinstance(x, str) and x.startswith("[") else []
        )
    for col in ["order_date", "delivery_date"]:
        if col in df.columns:
            df[col] = df[col].apply(excel_serial_to_date)
    return df


def generate_ai_brief(df: pd.DataFrame) -> str:
    cache_key = f"ai_brief_{len(df)}"
    if st.session_state.get("ai_brief_key") == cache_key:
        return st.session_state.get("ai_brief", "")

    try:
        import anthropic
        client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

        total      = len(df)
        clean      = int((df["status"] == "CLEAN").sum()) if "status" in df.columns else 0
        dupes      = int(df["status"].str.contains("DUPLICATE", na=False).sum()) if "status" in df.columns else 0
        mismatches = int(df["status"].str.contains("MISMATCH", na=False).sum()) if "status" in df.columns else 0
        flagged    = total - clean

        flagged_df   = df[df["status"].str.contains("DUPLICATE|MISMATCH", na=False)] if "status" in df.columns else pd.DataFrame()
        flag_details = ""
        for _, row in flagged_df.iterrows():
            flags     = row.get("flags", [])
            flag_text = "; ".join(flags) if isinstance(flags, list) else str(flags)
            flag_details += f"- PO {row.get('po_number','?')} (Supervisor {row.get('supervisor','?')}): {flag_text}\n"

        prompt = f"""You are writing a brief status update for a construction supply company's purchase order inbox.
Write 3-4 sentences in plain English. Be specific — mention PO numbers and supervisors when there are issues.
End with the single most important action needed right now. No bullet points, no headers.

Status:
- Total POs: {total} | Clean: {clean} | Flagged: {flagged}
- Duplicates: {dupes} | Vendor mismatches: {mismatches}

Issues:
{flag_details if flag_details else "None — all POs processed cleanly."}"""

        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=200,
            messages=[{"role": "user", "content": prompt}]
        )
        brief = response.content[0].text.strip()

    except Exception as e:
        total = len(df)
        clean = int((df["status"] == "CLEAN").sum()) if "status" in df.columns else 0
        brief = f"{total} purchase orders on file — {clean} processed cleanly. AI brief unavailable: check ANTHROPIC_API_KEY in .env."

    st.session_state["ai_brief"]     = brief
    st.session_state["ai_brief_key"] = cache_key
    return brief


def run_automation(dry_run: bool = False):
    cmd = ["python", "run_automation.py"]
    if dry_run:
        cmd.append("--dry-run")
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True,
            cwd=Path(__file__).parent, timeout=180
        )
        return result.stdout, result.stderr
    except subprocess.TimeoutExpired:
        return "", "Timed out after 3 minutes."
    except Exception as e:
        return "", str(e)


def run_demo_pipeline():
    try:
        result = subprocess.run(
            ["python", "run_demo.py"],
            capture_output=True, text=True,
            cwd=Path(__file__).parent, timeout=180
        )
        return result.stdout, result.stderr
    except subprocess.TimeoutExpired:
        return "", "Timed out after 3 minutes."
    except Exception as e:
        return "", str(e)


def inject_manual_files(uploaded_files) -> int:
    if not uploaded_files:
        return 0
    month    = datetime.now().strftime("%Y%m")
    dest_dir = Path(STORAGE_ROOT) / "inbox" / month / "Manual"
    dest_dir.mkdir(parents=True, exist_ok=True)
    count = 0
    for uf in uploaded_files:
        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = dest_dir / f"{ts}_{uf.name}"
        dest.write_bytes(uf.read())
        count += 1
    return count


def badge(status: str) -> str:
    if "DUPLICATE" in status: return '<span class="badge badge-dup">DUPLICATE</span>'
    if "MISMATCH"  in status: return '<span class="badge badge-mis">MISMATCH</span>'
    if "CLEAN"     in status: return '<span class="badge badge-cln">CLEAN</span>'
    return f'<span class="badge badge-err">{status}</span>'


def card_class(status: str) -> str:
    if "DUPLICATE" in status: return "alert-card dup"
    if "MISMATCH"  in status: return "alert-card mis"
    return "alert-card err"


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="iws-logo-mark">IWS</div>', unsafe_allow_html=True)
    st.markdown('<div class="iws-logo-sub">Purchase Order Automation</div>', unsafe_allow_html=True)
    st.markdown('<div class="red-bar"></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-head">Automation</div>', unsafe_allow_html=True)
    run_live = st.button("▶  RUN NOW", width='stretch', type="primary")
    run_dry  = st.button("◎  TEST RUN", width='stretch')

    if run_live:
        with st.spinner("Checking inbox and processing…"):
            out, err = run_automation(dry_run=False)
        st.session_state.pop("ai_brief", None)
        st.session_state.pop("ai_brief_key", None)
        st.success("Run complete.")
        if out:
            with st.expander("Output log"):
                st.code(out, language="text")
        if err:
            st.error(err)
        st.rerun()

    if run_dry:
        with st.spinner("Running test run…"):
            out, err = run_automation(dry_run=True)
        st.info("Test run — no files moved.")
        if out:
            with st.expander("Output log"):
                st.code(out, language="text")
        if err:
            st.warning(err)

    st.markdown('<div class="red-bar"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-head">Demo Mode</div>', unsafe_allow_html=True)
    st.markdown('<p style="font-size:0.8rem;color:rgba(255,255,255,0.6);margin-top:-4px;">Run pipeline against sample files — no Outlook needed.</p>', unsafe_allow_html=True)
    run_demo = st.button("◈  RUN DEMO", width='stretch')

    if run_demo:
        with st.spinner("Running demo pipeline…"):
            out, err = run_demo_pipeline()
        st.session_state.pop("ai_brief", None)
        st.session_state.pop("ai_brief_key", None)
        st.success("Demo run complete.")
        if out:
            with st.expander("Output log"):
                st.code(out, language="text")
        if err:
            st.warning(err)
        st.rerun()

    st.markdown('<div class="red-bar"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-head">Manual Submission</div>', unsafe_allow_html=True)
    st.markdown('<p style="font-size:0.8rem;color:rgba(255,255,255,0.6);margin-top:-4px;">Use when inbox access is unavailable.</p>', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Drop files",
        type=["xls", "xlsx", "xlsm", "pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )
    if uploaded:
        if st.button("⬆  SUBMIT FILES", width='stretch'):
            n = inject_manual_files(uploaded)
            st.success(f"{n} file(s) queued. Click Run Now to process.")

    st.markdown('<div class="red-bar"></div>', unsafe_allow_html=True)
    if st.button("↺  REFRESH", width='stretch'):
        st.session_state.pop("ai_brief", None)
        st.session_state.pop("ai_brief_key", None)
        st.rerun()
    st.caption(f"Updated {datetime.now().strftime('%H:%M:%S')}")


# ── Main ───────────────────────────────────────────────────────────────────────
st.markdown('<div class="page-title">Purchase Order Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="page-sub">Integrity Wall Systems &nbsp;·&nbsp; Automated Order Processing</div>', unsafe_allow_html=True)
st.markdown('<div class="red-bar"></div>', unsafe_allow_html=True)

df = load_tracking()

if df.empty:
    st.markdown("""
    <div style="text-align:center;padding:100px 0;">
        <div style="font-family:'Bebas Neue',sans-serif;font-size:1.2rem;letter-spacing:5px;color:rgba(255,255,255,0.5);">
            NO DATA YET
        </div>
        <div style="font-size:0.7rem;letter-spacing:3px;color:rgba(255,255,255,0.3);margin-top:8px;">
            RUN AUTOMATION OR SUBMIT FILES TO BEGIN
        </div>
    </div>""", unsafe_allow_html=True)
    st.stop()

# ── AI Brief ──────────────────────────────────────────────────────────────────
with st.spinner("Generating AI brief…"):
    brief = generate_ai_brief(df)

st.markdown(f"""
<div class="ai-brief-card">
    <div class="ai-brief-label">&#9889; AI INBOX BRIEF</div>
    <div class="ai-brief-text">{brief}</div>
</div>""", unsafe_allow_html=True)

# ── KPIs ──────────────────────────────────────────────────────────────────────
total      = len(df)
clean      = int((df["status"] == "CLEAN").sum()) if "status" in df.columns else 0
flagged    = total - clean
dupes      = int(df["status"].str.contains("DUPLICATE", na=False).sum()) if "status" in df.columns else 0
mismatches = int(df["status"].str.contains("MISMATCH",  na=False).sum()) if "status" in df.columns else 0

cols = st.columns(5)
for col, val, lbl, cls in [
    (cols[0], total,      "Total",    ""),
    (cols[1], clean,      "Clean",    "kpi-clean"),
    (cols[2], flagged,    "Flagged",  "kpi-flag"),
    (cols[3], dupes,      "Dupes",    "kpi-dup"),
    (cols[4], mismatches, "Mismatch", "kpi-mis"),
]:
    col.markdown(f"""
    <div class="kpi-wrap {cls}">
        <div class="kpi-n">{val}</div>
        <div class="kpi-l">{lbl}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Flagged items ─────────────────────────────────────────────────────────────
flagged_df = df[df["status"].str.contains("DUPLICATE|MISMATCH|VENDOR", na=False)] if "status" in df.columns else pd.DataFrame()

if not flagged_df.empty:
    with st.expander(f"⚠️  Flagged Items Requiring Attention  ({len(flagged_df)})", expanded=True):
        for _, row in flagged_df.iterrows():
            status   = str(row.get("status", ""))
            flags    = row.get("flags", [])
            po_num   = str(row.get("po_number", "")) if str(row.get("po_number", "")) != "nan" else "—"
            sup      = str(row.get("supervisor", "")) if str(row.get("supervisor", "")) != "nan" else "—"
            vendor   = str(row.get("vendor_on_form", "")) if str(row.get("vendor_on_form", "")) != "nan" else "—"
            src_file = str(row.get("source_file", ""))
            received = str(row.get("received_at", ""))[:16]
            flag_str = " &nbsp;·&nbsp; ".join(flags) if flags else ""

            st.markdown(f"""
            <div class="{card_class(status)}">
                {badge(status)}<strong>{po_num}</strong>
                <span class="alert-meta">&nbsp; Supervisor: {sup} &nbsp;·&nbsp; Vendor: {vendor} &nbsp;·&nbsp; {received} &nbsp;·&nbsp; {src_file}</span>
                <div class="alert-flag-text">{flag_str}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

# ── Filter bar ────────────────────────────────────────────────────────────────
st.markdown('<div class="section-head">PO Log</div>', unsafe_allow_html=True)

f1, f2, f3, f4 = st.columns([3, 1, 1, 1])
with f1:
    search = st.text_input("Search", "", placeholder="PO#, address, lot, vendor…", label_visibility="collapsed")
with f2:
    status_f = st.selectbox("Status", ["All", "CLEAN", "DUPLICATE", "VENDOR_MISMATCH", "PARSE_ERROR"], label_visibility="collapsed")
with f3:
    type_f = st.selectbox("Type", ["All"] + sorted(df["po_type"].dropna().unique().tolist()) if "po_type" in df.columns else ["All"], label_visibility="collapsed")
with f4:
    sup_f = st.selectbox("Supervisor", ["All"] + sorted(df["supervisor"].dropna().unique().tolist()) if "supervisor" in df.columns else ["All"], label_visibility="collapsed")

filtered = df.copy()
if search:
    filtered = filtered[filtered.apply(lambda r: search.lower() in str(r.values).lower(), axis=1)]
if status_f != "All" and "status" in filtered.columns:
    filtered = filtered[filtered["status"].str.contains(status_f, na=False)]
if type_f != "All" and "po_type" in filtered.columns:
    filtered = filtered[filtered["po_type"] == type_f]
if sup_f != "All" and "supervisor" in filtered.columns:
    filtered = filtered[filtered["supervisor"] == sup_f]

display_cols = [c for c in [
    "po_number", "po_type", "supervisor", "vendor_on_form",
    "order_date", "delivery_date", "lot", "address", "status", "logged_at"
] if c in filtered.columns]

st.caption(f"{len(filtered)} of {total} records")
st.dataframe(
    filtered[display_cols].sort_values("logged_at", ascending=False) if "logged_at" in filtered.columns else filtered[display_cols],
    use_container_width=True,
    hide_index=True,
)

# ── Detail view ───────────────────────────────────────────────────────────────
if "po_number" in filtered.columns and len(filtered) > 0:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-head">PO Detail</div>', unsafe_allow_html=True)
    valid_pos = [p for p in filtered["po_number"].dropna().unique().tolist() if str(p) != "nan"]
    if valid_pos:
        selected = st.selectbox("Select PO", ["Select…"] + sorted(valid_pos), label_visibility="collapsed")
        if selected != "Select…":
            row = filtered[filtered["po_number"] == selected].iloc[0]
            d1, d2, d3 = st.columns(3)
            with d1:
                st.markdown("**PO Info**")
                for k, v in [("Number","po_number"),("Type","po_type"),("Supervisor","supervisor"),("Category","category")]:
                    val = str(row.get(v,"")) if str(row.get(v,"")) != "nan" else "—"
                    st.write(f"**{k}:** {val}")
            with d2:
                st.markdown("**Dates & Location**")
                for k, v in [("Order Date","order_date"),("Delivery","delivery_date"),("Address","address"),("Lot","lot"),("Location","location")]:
                    val = str(row.get(v,"")) if str(row.get(v,"")) != "nan" else "—"
                    st.write(f"**{k}:** {val}")
            with d3:
                st.markdown("**Vendor & Status**")
                for k, v in [("Vendor on Form","vendor_on_form"),("Tract","tract"),("Release","release"),("Status","status")]:
                    val = str(row.get(v,"")) if str(row.get(v,"")) != "nan" else "—"
                    st.write(f"**{k}:** {val}")
                flags = row.get("flags", [])
                if flags:
                    for f in flags:
                        st.warning(f)
            items_json = row.get("po_items_json", "[]")
            try:
                items = json.loads(items_json) if isinstance(items_json, str) else []
                if items:
                    st.markdown("**Line Items**")
                    st.dataframe(pd.DataFrame(items), use_container_width=True, hide_index=True)
            except Exception:
                pass

# ── Export ────────────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown('<div class="section-head">Export</div>', unsafe_allow_html=True)
e1, e2 = st.columns(2)
with e1:
    csv = filtered[display_cols].to_csv(index=False).encode("utf-8")
    st.download_button("⬇  EXPORT LOG (CSV)", data=csv,
        file_name=f"po_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv", use_container_width=True)
with e2:
    if not flagged_df.empty:
        avail = [c for c in display_cols if c in flagged_df.columns]
        fcsv  = flagged_df[avail].to_csv(index=False).encode("utf-8")
        st.download_button("⬇  EXPORT FLAGGED (CSV)", data=fcsv,
            file_name=f"po_flagged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv", use_container_width=True)
