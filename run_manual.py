"""
run_manual.py
─────────────
Processes files that were manually uploaded via the dashboard drag-and-drop.
Scans data/inbox/YYYYMM/Manual/ and runs the full pipeline on any files found.
No Outlook connection required.

Usage:
    python run_manual.py              # live run — files processed, alerts sent
    python run_manual.py --dry-run    # process and log, no alert emails sent
"""

import sys
import os
import argparse
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

_SCRIPT_DIR  = Path(__file__).resolve().parent
STORAGE_ROOT = str(Path(os.getenv("STORAGE_ROOT", str(_SCRIPT_DIR / "data"))).resolve())

from agents.document_parser import parse_xls_po, parse_pdf_invoice
from agents.validator       import validate_po, validate_invoice, log_po_to_tracking
from agents.filer           import file_po, file_invoice, notify_if_flagged

SUPPORTED_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".pdf"}

NON_PO_ERRORS = [
    "not a po form",
    "unknown po type",
    "inventory",
    "report",
]


def scan_manual_inbox() -> list[dict]:
    """
    Scan all Manual subfolders under data/inbox/ and return a list of
    file dicts ready for processing.
    """
    inbox_root = Path(STORAGE_ROOT) / "inbox"
    if not inbox_root.exists():
        return []

    files = []
    for manual_dir in inbox_root.glob("*/Manual"):
        for f in sorted(manual_dir.iterdir()):
            if f.suffix.lower() in SUPPORTED_EXTENSIONS:
                files.append({
                    "filepath":  str(f),
                    "filename":  f.name,
                    "extension": f.suffix.lower(),
                })

    return files


def process_file(file_info: dict, dry_run: bool = False) -> dict | None:
    ext      = file_info["extension"]
    filepath = file_info["filepath"]
    filename = file_info["filename"]

    print(f"\n  → Processing: {filename} ({ext})")

    # Manually uploaded files have no email To: addresses
    to_addresses = []
    received_at  = datetime.now().isoformat()

    # ── Parse ─────────────────────────────────────────────────────────────────
    if ext in (".xls", ".xlsx", ".xlsm"):
        parsed   = parse_xls_po(filepath, to_addresses)
        doc_type = "PO"
    elif ext == ".pdf":
        parsed   = parse_pdf_invoice(filepath, {"sender": "manual", "subject": filename})
        doc_type = "Invoice"
    else:
        print(f"  [Skip] Unsupported: {ext}")
        return None

    if parsed.get("parse_error") and doc_type == "PO":
        print(f"  [Warning] Parse error: {parsed['parse_error']}")
        # Skip inventory reports, EOM sheets, and other non-PO files silently
        if any(kw in parsed["parse_error"].lower() for kw in NON_PO_ERRORS):
            print(f"  [Skip] Not a PO form — skipping tracking log: {filename}")
            return None

    # ── Validate ──────────────────────────────────────────────────────────────
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

    # ── File + Log ────────────────────────────────────────────────────────────
    if doc_type == "PO":
        dest = file_po(validated)
        log_po_to_tracking(validated, filename, received_at, to_addresses)
    else:
        dest = file_invoice(validated)

    validated["filed_to"] = dest

    if dry_run:
        print(f"  [Dry Run] Filed to: {dest} — alert email suppressed")
    else:
        email_meta = {
            "sender":       "manual upload",
            "subject":      filename,
            "to_addresses": [],
            "received_at":  received_at,
        }
        notify_if_flagged(validated, email_meta)

    # ── Remove from inbox after processing ───────────────────────────────────
    try:
        Path(filepath).unlink()
        print(f"  [Manual] Removed from inbox: {filename}")
    except Exception as e:
        print(f"  [Manual] Could not remove inbox file: {e}")

    return validated


def main():
    parser = argparse.ArgumentParser(description="IWS PO Automation — Manual File Processor")
    parser.add_argument("--dry-run", action="store_true",
                        help="Process and log files, but do not send alert emails.")
    args = parser.parse_args()

    dry_run = args.dry_run
    if dry_run:
        print("[Manual] DRY RUN MODE — files will be processed and logged, no alert emails sent.\n")

    start = datetime.now()
    print(f"[Manual] Started at {start.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"[Manual] Scanning: {Path(STORAGE_ROOT) / 'inbox'}\n")

    files = scan_manual_inbox()

    if not files:
        print("[Manual] No files found in inbox/Manual folders. Upload files via the dashboard first.")
        return

    print(f"[Manual] Found {len(files)} file(s) to process.")
    print(f"{'='*60}")

    all_results = []
    for file_info in files:
        result = process_file(file_info, dry_run=dry_run)
        if result:
            all_results.append(result)

    elapsed = (datetime.now() - start).total_seconds()
    clean   = sum(1 for r in all_results if r.get("validation_status") == "CLEAN")
    flagged = len(all_results) - clean

    print(f"\n{'='*60}")
    print(f"[Manual] Complete in {elapsed:.1f}s")
    print(f"  Processed : {len(all_results)}")
    print(f"  Clean     : {clean}")
    print(f"  Flagged   : {flagged}")
    if dry_run:
        print(f"  Mode      : DRY RUN — no alert emails sent")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()