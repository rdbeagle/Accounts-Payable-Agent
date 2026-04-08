"""
app.py — Streamlit Dashboard for Manual Review
Run with: streamlit run app.py
"""

import os
import json
import subprocess
import pandas as pd
import streamlit as st
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

STORAGE_ROOT = os.getenv("STORAGE_ROOT", "./data")
TRACKING_CSV = os.path.join(STORAGE_ROOT, "po_tracking.csv")

st.set_page_config(
    page_title="PO Automation Dashboard",
    page_icon="📦",
    layout="wide",
)

# ── Helpers ───────────────────────────────────────────────────────────────────

def load_tracking() -> pd.DataFrame:
    path = Path(TRACKING_CSV)
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_csv(path)
    # Parse flags from JSON string
    if "flags" in df.columns:
        df["flags"] = df["flags"].apply(
            lambda x: json.loads(x) if isinstance(x, str) and x.startswith("[") else []
        )
    return df


def status_color(status: str) -> str:
    if not status or status == "CLEAN":
        return "🟢"
    if "DUPLICATE" in status:
        return "🔴"
    if "MISMATCH" in status:
        return "🟠"
    if "ERROR" in status:
        return "⚫"
    return "🟡"


def trigger_run(dry_run: bool = False):
    cmd = ["python", "run_automation.py"]
    if dry_run:
        cmd.append("--dry-run")
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, cwd=Path(__file__).parent, timeout=120
        )
        return result.stdout, result.stderr
    except subprocess.TimeoutExpired:
        return "", "Run timed out after 120 seconds."
    except Exception as e:
        return "", str(e)


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://img.icons8.com/color/96/purchase-order.png", width=64)
    st.title("PO Automation")
    st.caption("Logistic Consultants")
    st.divider()

    st.subheader("⚙️ Manual Controls")
    col1, col2 = st.columns(2)
    with col1:
        run_live = st.button("▶ Run Now", use_container_width=True, type="primary")
    with col2:
        run_dry  = st.button("🧪 Dry Run", use_container_width=True)

    if run_live:
        with st.spinner("Checking inbox and processing emails…"):
            out, err = trigger_run(dry_run=False)
        st.success("Run complete.")
        if out:
            st.code(out, language="text")
        if err:
            st.error(err)

    if run_dry:
        with st.spinner("Running dry-run (no files moved)…"):
            out, err = trigger_run(dry_run=True)
        st.info("Dry run complete.")
        if out:
            st.code(out, language="text")
        if err:
            st.warning(err)

    st.divider()
    st.caption(f"Last refreshed: {datetime.now().strftime('%H:%M:%S')}")
    if st.button("🔄 Refresh Data"):
        st.rerun()

# ── Main content ──────────────────────────────────────────────────────────────
st.title("📦 Purchase Order Dashboard")

df = load_tracking()

if df.empty:
    st.info("No POs logged yet. Run the automation or check inbox credentials in `.env`.")
    st.stop()

# ── KPI row ───────────────────────────────────────────────────────────────────
total     = len(df)
clean     = (df["status"] == "CLEAN").sum() if "status" in df.columns else 0
flagged   = total - clean
dupes     = df["status"].str.contains("DUPLICATE", na=False).sum() if "status" in df.columns else 0
mismatches = df["status"].str.contains("MISMATCH", na=False).sum() if "status" in df.columns else 0

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total POs", total)
k2.metric("✅ Clean", clean)
k3.metric("⚠️ Flagged", flagged, delta=f"-{flagged}" if flagged else None,
          delta_color="inverse")
k4.metric("🔴 Duplicates", dupes)
k5.metric("🟠 Vendor Mismatches", mismatches)

st.divider()

# ── Filter bar ────────────────────────────────────────────────────────────────
col_f1, col_f2, col_f3, col_f4 = st.columns([2, 1, 1, 2])

with col_f1:
    search = st.text_input("🔍 Search PO#, address, vendor…", "")
with col_f2:
    status_filter = st.selectbox("Status", ["All", "CLEAN", "DUPLICATE", "VENDOR_MISMATCH", "PARSE_ERROR"])
with col_f3:
    type_filter = st.selectbox("PO Type", ["All"] + sorted(df["po_type"].dropna().unique().tolist()) if "po_type" in df.columns else ["All"])
with col_f4:
    supervisor_filter = st.selectbox("Supervisor", ["All"] + sorted(df["supervisor"].dropna().unique().tolist()) if "supervisor" in df.columns else ["All"])

filtered = df.copy()

if search:
    mask = filtered.apply(lambda row: search.lower() in str(row.values).lower(), axis=1)
    filtered = filtered[mask]
if status_filter != "All" and "status" in filtered.columns:
    filtered = filtered[filtered["status"].str.contains(status_filter, na=False)]
if type_filter != "All" and "po_type" in filtered.columns:
    filtered = filtered[filtered["po_type"] == type_filter]
if supervisor_filter != "All" and "supervisor" in filtered.columns:
    filtered = filtered[filtered["supervisor"] == supervisor_filter]

st.caption(f"Showing {len(filtered)} of {total} records")

# ── Flagged items first ───────────────────────────────────────────────────────
flagged_df = filtered[filtered["status"].str.contains(
    "DUPLICATE|MISMATCH|ERROR", na=False
)] if "status" in filtered.columns else pd.DataFrame()

if not flagged_df.empty:
    with st.expander(f"🚨 Flagged Items ({len(flagged_df)})", expanded=True):
        for _, row in flagged_df.iterrows():
            icon = status_color(row.get("status", ""))
            flags = row.get("flags", [])
            flag_str = " | ".join(flags) if flags else ""

            with st.container():
                c1, c2, c3 = st.columns([1, 2, 4])
                c1.markdown(f"**{icon} {row.get('po_number', 'N/A')}**")
                c2.markdown(f"`{row.get('status', '')}`")
                c3.markdown(f"⚠️ {flag_str}" if flag_str else "")
                st.caption(
                    f"Vendor: {row.get('vendor_on_form', '?')}  |  "
                    f"Supervisor: {row.get('supervisor', '?')}  |  "
                    f"Received: {row.get('received_at', '?')}  |  "
                    f"File: {row.get('source_file', '?')}"
                )
                st.divider()

# ── Full log table ─────────────────────────────────────────────────────────────
st.subheader("📋 Full PO Log")

display_cols = [c for c in [
    "po_number", "po_type", "supervisor", "vendor_on_form",
    "order_date", "delivery_date", "lot", "address",
    "status", "logged_at", "source_file"
] if c in filtered.columns]

display_df = filtered[display_cols].copy()
if "status" in display_df.columns:
    display_df["status"] = display_df["status"].apply(
        lambda s: f"{status_color(s)} {s}" if s else ""
    )

st.dataframe(
    display_df.sort_values("logged_at", ascending=False) if "logged_at" in display_df.columns else display_df,
    use_container_width=True,
    hide_index=True,
)

# ── Detail view ───────────────────────────────────────────────────────────────
st.subheader("🔎 PO Detail View")
if "po_number" in filtered.columns:
    po_options = ["Select a PO…"] + sorted(filtered["po_number"].dropna().unique().tolist())
    selected_po = st.selectbox("Select PO to inspect", po_options)

    if selected_po != "Select a PO…":
        row = filtered[filtered["po_number"] == selected_po].iloc[0]

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
            st.write(f"**Location:** {row.get('location')}")
        with d3:
            st.markdown("**Vendor & Status**")
            st.write(f"**Vendor on Form:** {row.get('vendor_on_form')}")
            icon = status_color(row.get("status", ""))
            st.write(f"**Status:** {icon} {row.get('status')}")
            flags = row.get("flags", [])
            if flags:
                st.markdown("**Flags:**")
                for f in flags:
                    st.warning(f)

        # Items
        items_json = row.get("po_items_json", "[]")
        try:
            items = json.loads(items_json) if isinstance(items_json, str) else []
            if items:
                st.markdown("**Line Items**")
                st.dataframe(pd.DataFrame(items), use_container_width=True, hide_index=True)
        except Exception:
            pass

# ── Export ────────────────────────────────────────────────────────────────────
st.divider()
st.subheader("📤 Export")
c_exp1, c_exp2 = st.columns(2)
with c_exp1:
    csv_data = filtered[display_cols].to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download Filtered Log (CSV)",
        data=csv_data,
        file_name=f"po_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv",
        use_container_width=True,
    )
with c_exp2:
    if not flagged_df.empty:
        flag_csv = flagged_df[display_cols].to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download Flagged Only (CSV)",
            data=flag_csv,
            file_name=f"po_flagged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
