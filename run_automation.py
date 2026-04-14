"""
run_automation.py
─────────────────
Main orchestration script. Run this on a schedule (Task Scheduler / cron)
to process new emails automatically.

Usage:
    python run_automation.py              # live run
    python run_automation.py --dry-run    # parse + validate, no files moved, no emails sent
"""

import argparse
from datetime import datetime
from pathlib import Path

# ── Agents ────────────────────────────────────────────────────────────────────
from agents.email_monitor   import fetch_new_emails, INVENTORY_PATTERNS
from agents.document_parser import parse_xls_po, parse_pdf_invoice
from agents.validator       import validate_po, validate_invoice, log_po_to_tracking
from agents.filer           import file_po, file_invoice, notify_if_flagged

import re

# Senders whose XLS files should not be parsed as POs
# (e.g. J&J Sand uses a completely different form format)
SKIP_PO_PARSE_SENDERS = {
    "jjsandtucson@gmail.com",
}


def _is_non_po_excel(filename: str) -> bool:
    """Return True if this Excel file is an inventory/report, not a PO form."""
    return any(re.search(p, filename) for p in INVENTORY_PATTERNS)


def process_email(email_meta: dict, dry_run: bool = False):
    """Process all attachments from a single email through all agents."""
    print(f"\n{'='*60}")
    print(f"[Orchestrator] Processing email: {email_meta['subject']}")
    print(f"  From   : {email_meta['sender']}")
    print(f"  To     : {', '.join(email_meta['to_addresses'])}")
    print(f"  Attachments: {len(email_meta['attachments'])}")

    sender       = email_meta.get("sender", "").lower()
    to_addresses = email_meta.get("to_addresses", [])
    results      = []

    for attachment in email_meta["attachments"]:
        ext      = attachment["extension"]
        filepath = attachment["filepath"]
        filename = attachment["filename"]

        print(f"\n  → Processing: {filename} ({ext})")

        # ── AGENT 2: Parse ────────────────────────────────────────────────────
        if ext in (".xls", ".xlsx", ".xlsm"):

            # Skip inventory/report Excel files — already routed to Inventory/
            if _is_non_po_excel(filename):
                print(f"  [Skip] Inventory/report file — not a PO form")
                continue

            # Skip known non-PO senders (J&J Sand, etc.)
            if any(skip in sender for skip in SKIP_PO_PARSE_SENDERS):
                print(f"  [Skip] Sender not a PO source: {sender}")
                continue

            parsed   = parse_xls_po(filepath, to_addresses)
            doc_type = "PO"

        elif ext == ".pdf":
            parsed   = parse_pdf_invoice(filepath, email_meta)
            doc_type = "Invoice"
        else:
            print(f"  [Skip] Unsupported extension: {ext}")
            continue

        if parsed.get("parse_error") and doc_type == "PO":
            print(f"  [Warning] Parse error: {parsed['parse_error']}")

        # ── AGENT 3: Validate ─────────────────────────────────────────────────
        if doc_type == "PO":
            validated = validate_po(parsed)
            print(f"  Validation: {validated['validation_status']}")
            if validated["validation_flags"]:
                for f in validated["validation_flags"]:
                    print(f"    ⚠  {f}")
        else:
            validated = validate_invoice(parsed)
            print(f"  Invoice validation: {validated['validation_status']}")

        # ── AGENT 4: File + Notify ────────────────────────────────────────────
        if not dry_run:
            if doc_type == "PO":
                dest = file_po(validated)
                # Only log to tracking if it parsed as a real PO with a number
                po_num = validated.get("po_number", "")
                if po_num and po_num not in ("UNKNOWN", "", "None"):
                    log_po_to_tracking(validated, filename, email_meta["received_at"])
                else:
                    print(f"  [Skip log] No PO number — not logged to tracking")
            else:
                dest = file_invoice(validated)

            validated["filed_to"] = dest
            notify_if_flagged(validated, email_meta)
        else:
            print(f"  [Dry Run] Would file to: {'flagged' if validated.get('validation_flags') else 'processed'}/")

        results.append(validated)

    return results


def main():
    parser = argparse.ArgumentParser(description="PO Automation — Email Processing Agent")
    parser.add_argument("--dry-run", action="store_true",
                        help="Parse and validate without moving files or sending emails.")
    args = parser.parse_args()

    dry_run = args.dry_run
    if dry_run:
        print("[Orchestrator] DRY RUN MODE — no files will be moved or emails sent.\n")

    start = datetime.now()
    print(f"[Orchestrator] Run started at {start.strftime('%Y-%m-%d %H:%M:%S')}")

    # ── AGENT 1: Fetch emails ────────────────────────────────────────────────
    emails = fetch_new_emails(mark_seen=not dry_run)

    if not emails:
        print("[Orchestrator] No new emails with attachments. Exiting.")
        return

    all_results = []
    for email_meta in emails:
        results = process_email(email_meta, dry_run=dry_run)
        all_results.extend(results)

    # ── Summary ──────────────────────────────────────────────────────────────
    elapsed  = (datetime.now() - start).total_seconds()
    clean    = sum(1 for r in all_results if r.get("validation_status") == "CLEAN")
    flagged  = len(all_results) - clean

    print(f"\n{'='*60}")
    print(f"[Orchestrator] Run complete in {elapsed:.1f}s")
    print(f"  Processed : {len(all_results)} document(s)")
    print(f"  Clean     : {clean}")
    print(f"  Flagged   : {flagged}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
