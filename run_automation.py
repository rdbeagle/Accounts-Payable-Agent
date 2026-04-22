"""
run_automation.py
─────────────────
Main orchestration script. Run this on a schedule (Task Scheduler / cron)
to process new emails automatically.

Usage:
    python run_automation.py              # live run
    python run_automation.py --dry-run    # file + log, but no emails sent, no messages marked read
"""

import sys
import argparse
from datetime import datetime

# ── Agents ────────────────────────────────────────────────────────────────────
from agents.email_monitor   import fetch_new_emails
from agents.document_parser import parse_xls_po, parse_pdf_invoice
from agents.validator       import validate_po, validate_invoice, log_po_to_tracking
from agents.filer           import file_po, file_invoice, notify_if_flagged


def process_email(email_meta: dict, dry_run: bool = False):
    """Process all attachments from a single email through all agents."""
    print(f"\n{'='*60}")
    print(f"[Orchestrator] Processing email: {email_meta['subject']}")
    print(f"  From   : {email_meta['sender']}")
    print(f"  To     : {', '.join(email_meta['to_addresses'])}")
    print(f"  Attachments: {len(email_meta['attachments'])}")

    to_addresses = email_meta.get("to_addresses", [])
    results = []

    for attachment in email_meta["attachments"]:
        ext      = attachment["extension"]
        filepath = attachment["filepath"]
        filename = attachment["filename"]

        print(f"\n  → Processing: {filename} ({ext})")

        # ── AGENT 2: Parse ────────────────────────────────────────────────────
        if ext in (".xls", ".xlsx", ".xlsm"):
            parsed = parse_xls_po(filepath, to_addresses)
            doc_type = "PO"
        elif ext == ".pdf":
            parsed = parse_pdf_invoice(filepath, email_meta)
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
        # Always file and log — dry run only skips emails and marking read
        if doc_type == "PO":
            dest = file_po(validated)
            log_po_to_tracking(validated, filename, email_meta["received_at"], to_addresses)
        else:
            dest = file_invoice(validated)

        validated["filed_to"] = dest

        if dry_run:
            print(f"  [Dry Run] Filed to: {dest} — alert email suppressed")
        else:
            notify_if_flagged(validated, email_meta)

        results.append(validated)

    return results


def main():
    parser = argparse.ArgumentParser(description="PO Automation — Email Processing Agent")
    parser.add_argument("--dry-run", action="store_true",
                        help="File and log normally, but skip sending alert emails and marking messages as read.")
    args = parser.parse_args()

    dry_run = args.dry_run
    if dry_run:
        print("[Orchestrator] DRY RUN MODE — files will be saved and logged, but no alert emails will be sent and Outlook messages will not be marked as read.\n")

    start = datetime.now()
    print(f"[Orchestrator] Run started at {start.strftime('%Y-%m-%d %H:%M:%S')}")

    # ── AGENT 1: Fetch emails ─────────────────────────────────────────────────
    # mark_seen=False in dry run so emails stay unread and can be reprocessed
    emails = fetch_new_emails(mark_seen=not dry_run)

    if not emails:
        print("[Orchestrator] No new emails with attachments. Exiting.")
        return

    all_results = []
    for email_meta in emails:
        results = process_email(email_meta, dry_run=dry_run)
        all_results.extend(results)

    # ── Summary ──────────────────────────────────────────────────────────────
    elapsed = (datetime.now() - start).total_seconds()
    clean   = sum(1 for r in all_results if r.get("validation_status") == "CLEAN")
    flagged = len(all_results) - clean

    print(f"\n{'='*60}")
    print(f"[Orchestrator] Run complete in {elapsed:.1f}s")
    print(f"  Processed : {len(all_results)} document(s)")
    print(f"  Clean     : {clean}")
    print(f"  Flagged   : {flagged}")
    if dry_run:
        print(f"  Mode      : DRY RUN — no alert emails sent, messages left unread")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()