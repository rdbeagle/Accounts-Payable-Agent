"""
run_demo.py
───────────
Demo mode — runs the full pipeline against local sample files
instead of connecting to Outlook. Use this for presentations
and grader submissions where inbox access is unavailable.

Usage:
    python run_demo.py              # full demo run
    python run_demo.py --dry-run    # parse and validate only

Setup:
    1. Place anonymized PO files (.xls) in demo/attachments/pos/
    2. Place sample invoice PDFs in demo/attachments/invoices/
    3. Edit demo/emails.json to describe the sample emails
    4. Run: python run_demo.py
"""

import argparse
import json
import shutil
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
import os

load_dotenv()

STORAGE_ROOT = os.getenv("STORAGE_ROOT", "./po-automation/data")

from agents.document_parser import parse_xls_po, parse_pdf_invoice
from agents.validator       import validate_po, validate_invoice, log_po_to_tracking
from agents.filer           import file_po, file_invoice, notify_if_flagged


DEMO_DIR = Path(__file__).parent / "demo"


def load_demo_emails() -> list[dict]:
    """
    Load simulated email metadata from demo/emails.json and
    pair each with its attachment file from demo/attachments/.
    """
    emails_file = DEMO_DIR / "emails.json"
    if not emails_file.exists():
        print(f"[Demo] No emails.json found at {emails_file}")
        print("[Demo] Run: python run_demo.py --setup  to create sample files")
        return []

    with open(emails_file, encoding="utf-8") as f:
        email_defs = json.load(f)

    # Copy attachments to a temp inbox folder so the pipeline can find them
    inbox_dir = Path(STORAGE_ROOT) / "inbox" / datetime.now().strftime("%Y%m") / "Demo"
    inbox_dir.mkdir(parents=True, exist_ok=True)

    results = []
    for em in email_defs:
        attachments = []
        for att in em.get("attachments", []):
            src = DEMO_DIR / "attachments" / att["filename"]
            if not src.exists():
                print(f"[Demo] Attachment not found: {src}")
                continue
            ts        = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest_name = f"{ts}_{att['filename']}"
            dest      = inbox_dir / dest_name
            shutil.copy2(src, dest)
            attachments.append({
                "filename":  att["filename"],
                "filepath":  str(dest),
                "extension": Path(att["filename"]).suffix.lower(),
            })

        if attachments:
            results.append({
                "uid":          em.get("uid", f"demo_{len(results)}"),
                "subject":      em.get("subject", "Demo Email"),
                "sender":       em.get("sender", "demo@example.com"),
                "to_addresses": em.get("to_addresses", []),
                "received_at":  em.get("received_at", datetime.now().isoformat()),
                "vendor":       em.get("vendor", "Demo"),
                "attachments":  attachments,
            })

    print(f"[Demo] Loaded {len(results)} demo email(s) with attachments")
    return results


def process_email(email_meta: dict, dry_run: bool = False):
    print(f"\n{'='*60}")
    print(f"[Demo] Processing: {email_meta['subject']}")
    print(f"  From : {email_meta['sender']}")
    print(f"  To   : {', '.join(email_meta['to_addresses'])}")

    to_addresses = email_meta.get("to_addresses", [])
    results      = []

    for attachment in email_meta["attachments"]:
        ext      = attachment["extension"]
        filepath = attachment["filepath"]
        filename = attachment["filename"]

        print(f"\n  → {filename} ({ext})")

        if ext in (".xls", ".xlsx", ".xlsm"):
            parsed   = parse_xls_po(filepath, to_addresses)
            doc_type = "PO"
        elif ext == ".pdf":
            parsed   = parse_pdf_invoice(filepath, email_meta)
            doc_type = "Invoice"
        else:
            print(f"  [Skip] {ext}")
            continue

        if parsed.get("parse_error") and doc_type == "PO":
            print(f"  [Warning] {parsed['parse_error']}")

        if doc_type == "PO":
            validated = validate_po(parsed)
            print(f"  Validation: {validated['validation_status']}")
            for f in validated.get("validation_flags", []):
                print(f"    ⚠  {f}")
        else:
            validated = validate_invoice(parsed)
            print(f"  Invoice: {validated['validation_status']}")
            for f in validated.get("validation_flags", []):
                print(f"    ⚠  {f}")

        if not dry_run:
            if doc_type == "PO":
                dest   = file_po(validated)
                po_num = validated.get("po_number", "")
                if po_num and po_num not in ("UNKNOWN", "", "None"):
                    log_po_to_tracking(validated, filename, email_meta["received_at"])
            else:
                dest = file_invoice(validated)
            validated["filed_to"] = dest
            notify_if_flagged(validated, email_meta)
        else:
            status = validated.get("validation_flags", [])
            print(f"  [Dry Run] → {'flagged' if status else 'processed'}/")

        results.append(validated)

    return results


def create_sample_files():
    """Create the demo folder structure and a sample emails.json template."""
    (DEMO_DIR / "attachments").mkdir(parents=True, exist_ok=True)

    sample = [
        {
            "uid": "demo_001",
            "subject": "LH ORDER LOT 94 - DELIVERY TOMORROW",
            "sender": "blakek@integrityllctuc.com",
            "to_addresses": ["tucsonstucco@rwc.org", "orders@logisticconsultants.com"],
            "received_at": "2026-04-14T09:00:00",
            "vendor": "RWC",
            "attachments": [{"filename": "9001BK_LH_LOT94.xls"}]
        },
        {
            "uid": "demo_002",
            "subject": "ST ORDER LOT 30 SBR - WILLCALL",
            "sender": "blakek@integrityllctuc.com",
            "to_addresses": [
                "adrian.carino@lwsupply.com",
                "orders@logisticconsultants.com"
            ],
            "received_at": "2026-04-14T09:15:00",
            "vendor": "LW",
            "attachments": [{"filename": "9002BK_ST_LOT30.xls"}]
        },
        {
            "uid": "demo_003",
            "subject": "LH ORDER - VENDOR MISMATCH EXAMPLE",
            "sender": "blakek@integrityllctuc.com",
            "to_addresses": [
                "adrian.carino@lwsupply.com",
                "orders@logisticconsultants.com"
            ],
            "received_at": "2026-04-14T09:30:00",
            "vendor": "LW",
            "attachments": [{"filename": "9003BK_LH_MISMATCH.xls"}]
        },
        {
            "uid": "demo_004",
            "subject": "RE: LH ORDER LOT 94 - DUPLICATE",
            "sender": "blakek@integrityllctuc.com",
            "to_addresses": ["tucsonstucco@rwc.org", "orders@logisticconsultants.com"],
            "received_at": "2026-04-14T10:00:00",
            "vendor": "RWC",
            "attachments": [{"filename": "9001BK_LH_LOT94.xls"}]
        },
        {
            "uid": "demo_005",
            "subject": "Invoice from RWC Supply",
            "sender": "donotreply@rwc.org",
            "to_addresses": ["orders@logisticconsultants.com"],
            "received_at": "2026-04-14T11:00:00",
            "vendor": "RWC",
            "attachments": [{"filename": "RWC_Invoice_demo.pdf"}]
        },
    ]

    out = DEMO_DIR / "emails.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(sample, f, indent=2)

    readme = DEMO_DIR / "README.md"
    readme.write_text("""# Demo Mode — IWS PO Automation

## Setup
1. Place anonymized PO files in `demo/attachments/`:
   - `9001BK_LH_LOT94.xls`   — clean Lath PO (RWC)
   - `9002BK_ST_LOT30.xls`   — clean Stucco PO (L&W)
   - `9003BK_LH_MISMATCH.xls` — Lath PO with vendor mismatch (form says RWC, sent to L&W)
   - `RWC_Invoice_demo.pdf`   — sample invoice referencing PO# 9001BK

2. Run the demo:
   ```
   python run_demo.py
   streamlit run app.py
   ```

## What it demonstrates
- ✅ XLS PO parsing (Lath and Stucco form types)
- ✅ Vendor mismatch detection (9003BK sent to wrong vendor)
- ✅ Duplicate PO detection (9001BK submitted twice)
- ✅ Invoice-to-PO matching (RWC invoice matched to 9001BK)
- ✅ Claude-generated alert emails
- ✅ AI dashboard brief
- ✅ File organization by vendor/month
""")

    print(f"[Demo] Created demo folder at {DEMO_DIR}")
    print(f"[Demo] Add your anonymized files to {DEMO_DIR / 'attachments'}/")
    print(f"[Demo] Then run: python run_demo.py")


def main():
    parser = argparse.ArgumentParser(description="IWS PO Automation — Demo Mode")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--setup",   action="store_true",
                        help="Create demo folder structure and sample emails.json")
    args = parser.parse_args()

    if args.setup:
        create_sample_files()
        return

    if args.dry_run:
        print("[Demo] DRY RUN — no files moved.\n")

    start = datetime.now()
    print(f"[Demo] Started at {start.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"[Demo] Loading emails from {DEMO_DIR}/emails.json\n")

    emails = load_demo_emails()
    if not emails:
        print("[Demo] No emails to process. Run: python run_demo.py --setup")
        return

    all_results = []
    for email_meta in emails:
        results = process_email(email_meta, dry_run=args.dry_run)
        all_results.extend(results)

    elapsed = (datetime.now() - start).total_seconds()
    clean   = sum(1 for r in all_results if r.get("validation_status") == "CLEAN")
    flagged = len(all_results) - clean

    print(f"\n{'='*60}")
    print(f"[Demo] Complete in {elapsed:.1f}s")
    print(f"  Processed : {len(all_results)}")
    print(f"  Clean     : {clean}")
    print(f"  Flagged   : {flagged}")
    print(f"{'='*60}")
    print(f"\nNow run: streamlit run app.py")


if __name__ == "__main__":
    main()
