"""
app.py — Integrity Wall Systems PO Automation Dashboard
Run with: streamlit run app.py
"""

import os
import json
import subprocess
import shutil
import pandas as pd
import streamlit as st
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

STORAGE_ROOT = os.getenv("STORAGE_ROOT", "./data")
TRACKING_CSV = os.path.join(STORAGE_ROOT, "po_tracking.csv")

st.set_page_config(
    page_title="IWS — PO Automation",
    page_icon="🏗️",
    layout="wide",
)

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500&display=swap');

/* Flag background overlay */
[data-testid="stAppViewContainer"] {
    background-image:
        linear-gradient(135deg, rgba(12,18,42,0.93) 0%, rgba(20,40,80,0.90) 50%, rgba(12,18,42,0.95) 100%),
        repeating-linear-gradient(
            -55deg,
            transparent,
            transparent 60px,
            rgba(180,30,30,0.07) 60px,
            rgba(180,30,30,0.07) 62px
        ),
        repeating-linear-gradient(
            35deg,
            transparent,
            transparent 60px,
            rgba(255,255,255,0.04) 60px,
            rgba(255,255,255,0.04) 62px
        );
    background-attachment: fixed;
}

[data-testid="stHeader"] { background: transparent !important; }
[data-testid="stSidebar"] {
    background: rgba(8,14,36,0.97) !important;
    border-right: 1px solid rgba(180,30,30,0.3) !important;
}

/* Typography */
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif !important; }

h1, h2, h3 { font-family: 'Bebas Neue', sans-serif !important; letter-spacing: 2px; }

/* Main title */
.iws-title {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 3.2rem;
    letter-spacing: 6px;
    color: #ffffff;
    text-shadow: 0 0 40px rgba(220,40,40,0.5);
    margin: 0;
    line-height: 1;
}
.iws-subtitle {
    font-family: 'DM Sans', sans-serif;
    font-size: 0.75rem;
    letter-spacing: 4px;
    text-transform: uppercase;
    color: rgba(255,255,255,0.4);
    margin-top: 4px;
}
.iws-divider {
    height: 3px;
    background: linear-gradient(90deg, #B41E1E, #E63232, transparent);
    border: none;
    margin: 8px 0 24px 0;
}

/* KPI cards */
.kpi-card {
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 8px;
    padding: 20px 24px;
    text-align: center;
}
.kpi-number {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 2.8rem;
    line-height: 1;
    color: #fff;
}
.kpi-label {
    font-size: 0.68rem;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: rgba(255,255,255,0.4);
    margin-top: 4px;
}
.kpi-clean .kpi-number  { color: #4ade80; }
.kpi-flag .kpi-number   { color: #f87171; }
.kpi-dupe .kpi-number   { color: #E63232; }
.kpi-miss .kpi-number   { color: #fb923c; }

/* Status badges */
.badge-clean   { background:#14532d; color:#4ade80; padding:2px 10px; border-radius:4px; font-size:0.72rem; letter-spacing:1px; }
.badge-flag    { background:#450a0a; color:#f87171; padding:2px 10px; border-radius:4px; font-size:0.72rem; letter-spacing:1px; }
.badge-dup     { background:#7f1d1d; color:#fca5a5; padding:2px 10px; border-radius:4px; font-size:0.72rem; letter-spacing:1px; }
.badge-miss    { background:#431407; color:#fb923c; padding:2px 10px; border-radius:4px; font-size:0.72rem; letter-spacing:1px; }

/* Run button */
.stButton > button {
    background: linear-gradient(135deg, #B41E1E, #E63232) !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'Bebas Neue', sans-serif !important;
    letter-spacing: 3px !important;
    font-size: 1rem !important;
    padding: 12px 24px !important;
    transition: all 0.2s !important;
    box-shadow: 0 4px 20px rgba(180,30,30,0.3) !important;
}
.stButton > button:hover {
    box-shadow: 0 4px 30px rgba(230,50,50,0.5) !important;
    transform: translateY(-1px) !important;
}

/* Drop zone styling */
[data-testid="stFileUploader"] {
    background: rgba(255,255,255,0.03) !important;
    border: 2px dashed rgba(180,30,30,0.4) !important;
    border-radius: 8px !important;
}

/* Table */
[data-testid="stDataFrame"] {
    background: rgba(255,255,255,0.02) !important;
    border: 1px solid rgba(255,255,255,0.06) !important;
    border-radius: 8px !important;
}

/* Section headers */
.section-label {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 1.1rem;
    letter-spacing: 4px;
    color: rgba(255,255,255,0.5);
    text-transform: uppercase;
    margin-bottom: 12px;
    padding-bottom: 6px;
    border-bottom: 1px solid rgba(255,255,255,0.07);
}

/* Alert banner */
.alert-banner {
    background: rgba(180,30,30,0.15);
    border: 1px solid rgba(180,30,30,0.4);
    border-left: 4px solid #E63232;
    border-radius: 6px;
    padding: 12px 16px;
    margin-bottom: 10px;
    font-size: 0.85rem;
    color: rgba(255,255,255,0.85);
}
</style>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────
def load_tracking() -> pd.DataFrame:
    path = Path(TRACKING_CSV)
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_csv(path)
    if "flags" in df.columns:
        df["flags"] = df["flags"].apply(
            lambda x: json.loads(x) if isinstance(x, str) and x.startswith("[") else []
        )
    return df


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


def inject_manual_files(uploaded_files) -> int:
    """Save manually uploaded files into the inbox for processing."""
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


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<p class="iws-title">IWS</p>', unsafe_allow_html=True)
    st.markdown('<p class="iws-subtitle">Purchase Order Automation</p>', unsafe_allow_html=True)
    st.markdown('<hr class="iws-divider">', unsafe_allow_html=True)

    st.markdown('<p class="section-label">Automation Controls</p>', unsafe_allow_html=True)

    run_live = st.button("▶  RUN NOW", use_container_width=True, type="primary")
    run_dry  = st.button("◎  DRY RUN", use_container_width=True)

    if run_live:
        with st.spinner("Processing inbox…"):
            out, err = run_automation(dry_run=False)
        st.success("Run complete.")
        if out:
            with st.expander("Output"):
                st.code(out, language="text")
        if err:
            st.error(err)
        st.rerun()

    if run_dry:
        with st.spinner("Running dry run…"):
            out, err = run_automation(dry_run=True)
        st.info("Dry run complete — no files moved.")
        if out:
            with st.expander("Output"):
                st.code(out, language="text")
        if err:
            st.warning(err)

    st.markdown('<hr class="iws-divider">', unsafe_allow_html=True)
    st.markdown('<p class="section-label">Manual File Submission</p>', unsafe_allow_html=True)
    st.caption("Upload POs or invoices when inbox access is unavailable.")

    uploaded = st.file_uploader(
        "Drop files here",
        type=["xls", "xlsx", "xlsm", "pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if uploaded:
        if st.button("⬆  SUBMIT FILES", use_container_width=True):
            n = inject_manual_files(uploaded)
            st.success(f"{n} file(s) queued. Run automation to process.")

    st.markdown('<hr class="iws-divider">', unsafe_allow_html=True)
    if st.button("↺  REFRESH", use_container_width=True):
        st.rerun()
    st.caption(f"Last refreshed: {datetime.now().strftime('%H:%M:%S')}")


# ── Main content ──────────────────────────────────────────────────────────────
st.markdown('<p class="iws-title">Purchase Order Dashboard</p>', unsafe_allow_html=True)
st.markdown('<p class="iws-subtitle">Integrity Wall Systems — Automated Order Processing</p>', unsafe_allow_html=True)
st.markdown('<hr class="iws-divider">', unsafe_allow_html=True)

df = load_tracking()

if df.empty:
    st.markdown("""
    <div style="text-align:center;padding:80px 0;color:rgba(255,255,255,0.3);">
        <div style="font-family:'Bebas Neue',sans-serif;font-size:1.4rem;letter-spacing:4px;">
            NO DATA YET
        </div>
        <div style="font-size:0.8rem;margin-top:8px;letter-spacing:2px;">
            RUN AUTOMATION OR SUBMIT FILES TO BEGIN
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── KPIs ──────────────────────────────────────────────────────────────────────
total    = len(df)
clean    = int((df["status"] == "CLEAN").sum()) if "status" in df.columns else 0
flagged  = total - clean
dupes    = int(df["status"].str.contains("DUPLICATE", na=False).sum()) if "status" in df.columns else 0
mismatches = int(df["status"].str.contains("MISMATCH", na=False).sum()) if "status" in df.columns else 0

k1, k2, k3, k4, k5 = st.columns(5)
for col, val, label, cls in [
    (k1, total,      "Total POs",         ""),
    (k2, clean,      "Clean",             "kpi-clean"),
    (k3, flagged,    "Flagged",           "kpi-flag"),
    (k4, dupes,      "Duplicates",        "kpi-dupe"),
    (k5, mismatches, "Vendor Mismatches", "kpi-miss"),
]:
    col.markdown(f"""
    <div class="kpi-card {cls}">
        <div class="kpi-number">{val}</div>
        <div class="kpi-label">{label}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Flagged alerts ────────────────────────────────────────────────────────────
flagged_df = df[df["status"].str.contains("DUPLICATE|MISMATCH|ERROR", na=False)] if "status" in df.columns else pd.DataFrame()

if not flagged_df.empty:
    st.markdown('<p class="section-label">⚠ Flagged Items Requiring Attention</p>', unsafe_allow_html=True)
    for _, row in flagged_df.iterrows():
        status = row.get("status", "")
        flags  = row.get("flags", [])
        flag_str = " &nbsp;|&nbsp; ".join(flags) if flags else ""
        badge_cls = "badge-dup" if "DUPLICATE" in status else "badge-miss" if "MISMATCH" in status else "badge-flag"
        st.markdown(f"""
        <div class="alert-banner">
            <span class="{badge_cls}">{status}</span>
            &nbsp;&nbsp;<strong>{row.get('po_number','?')}</strong>
            &nbsp;·&nbsp; Supervisor: {row.get('supervisor','?')}
            &nbsp;·&nbsp; Vendor: {row.get('vendor_on_form','?')}
            &nbsp;·&nbsp; {row.get('received_at','')}<br>
            <span style="color:rgba(255,255,255,0.5);font-size:0.8rem;">{flag_str}</span>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

# ── Filters ───────────────────────────────────────────────────────────────────
st.markdown('<p class="section-label">PO Log</p>', unsafe_allow_html=True)

fc1, fc2, fc3, fc4 = st.columns([3, 1, 1, 1])
with fc1:
    search = st.text_input("Search", "", placeholder="PO#, address, vendor…", label_visibility="collapsed")
with fc2:
    status_filter = st.selectbox("Status", ["All", "CLEAN", "DUPLICATE", "VENDOR_MISMATCH", "PARSE_ERROR"], label_visibility="collapsed")
with fc3:
    type_filter = st.selectbox("Type", ["All"] + sorted(df["po_type"].dropna().unique().tolist()) if "po_type" in df.columns else ["All"], label_visibility="collapsed")
with fc4:
    sup_filter = st.selectbox("Supervisor", ["All"] + sorted(df["supervisor"].dropna().unique().tolist()) if "supervisor" in df.columns else ["All"], label_visibility="collapsed")

filtered = df.copy()
if search:
    filtered = filtered[filtered.apply(lambda r: search.lower() in str(r.values).lower(), axis=1)]
if status_filter != "All" and "status" in filtered.columns:
    filtered = filtered[filtered["status"].str.contains(status_filter, na=False)]
if type_filter != "All" and "po_type" in filtered.columns:
    filtered = filtered[filtered["po_type"] == type_filter]
if sup_filter != "All" and "supervisor" in filtered.columns:
    filtered = filtered[filtered["supervisor"] == sup_filter]

display_cols = [c for c in ["po_number","po_type","supervisor","vendor_on_form","order_date","delivery_date","lot","address","status","logged_at"] if c in filtered.columns]

st.caption(f"{len(filtered)} of {total} records")
st.dataframe(
    filtered[display_cols].sort_values("logged_at", ascending=False) if "logged_at" in filtered.columns else filtered[display_cols],
    use_container_width=True,
    hide_index=True,
)

# ── Detail view ───────────────────────────────────────────────────────────────
if "po_number" in filtered.columns and len(filtered) > 0:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<p class="section-label">PO Detail</p>', unsafe_allow_html=True)
    options = ["Select…"] + sorted(filtered["po_number"].dropna().unique().tolist())
    selected = st.selectbox("Select PO", options, label_visibility="collapsed")

    if selected != "Select…":
        row = filtered[filtered["po_number"] == selected].iloc[0]
        d1, d2, d3 = st.columns(3)
        with d1:
            st.markdown("**PO Info**")
            st.write(f"**Number:** {row.get('po_number')}")
            st.write(f"**Type:** {row.get('po_type')}")
            st.write(f"**Supervisor:** {row.get('supervisor')}")
            st.write(f"**Category:** {row.get('category')}")
        with d2:
            st.markdown("**Dates & Location**")
            st.write(f"**Order Date:** {row.get('order_date')}")
            st.write(f"**Delivery:** {row.get('delivery_date')}")
            st.write(f"**Address:** {row.get('address')}")
            st.write(f"**Lot:** {row.get('lot')}")
        with d3:
            st.markdown("**Vendor & Status**")
            st.write(f"**Vendor on Form:** {row.get('vendor_on_form')}")
            st.write(f"**Status:** {row.get('status')}")
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
c1, c2 = st.columns(2)
with c1:
    csv = filtered[display_cols].to_csv(index=False).encode("utf-8")
    st.download_button("⬇  EXPORT LOG (CSV)", data=csv,
        file_name=f"po_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv", use_container_width=True)
with c2:
    if not flagged_df.empty:
        fcsv = flagged_df[display_cols].to_csv(index=False).encode("utf-8")
        st.download_button("⬇  EXPORT FLAGGED (CSV)", data=fcsv,
            file_name=f"po_flagged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv", use_container_width=True)
