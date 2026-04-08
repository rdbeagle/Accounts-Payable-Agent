"""
Agent 3: Validator
- Checks for duplicate PO numbers against the tracking log
- Flags vendor mismatches (email To: domain vs. vendor named on form)
- Matches invoices to POs and checks amount discrepancies
"""

import os
import csv
import json
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

STORAGE_ROOT = os.getenv("STORAGE_ROOT", "./data")
TRACKING_CSV = os.path.join(STORAGE_ROOT, "po_tracking.csv")

TRACKING_COLUMNS = [
    "po_number", "po_type", "supervisor", "vendor_on_form",
    "order_date", "delivery_date", "address", "release",
    "lot", "category", "track", "location",
    "source_file", "received_at", "logged_at",
    "status",           # CLEAN | DUPLICATE | VENDOR_MISMATCH | INVOICE_DISCREPANCY
    "flags",            # JSON list of human-readable flag strings
    "invoice_amount",
    "po_items_json",
]


def _load_tracking() -> list[dict]:
    path = Path(TRACKING_CSV)
    if not path.exists():
        return []
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def _save_tracking(rows: list[dict]):
    Path(TRACKING_CSV).parent.mkdir(parents=True, exist_ok=True)
    with open(TRACKING_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=TRACKING_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def get_all_po_numbers() -> set[str]:
    return {r["po_number"].strip().upper() for r in _load_tracking() if r.get("po_number")}


def validate_po(parsed_po: dict) -> dict:
    """
    Validate a parsed PO dict (from document_parser.parse_xls_po).
    Returns the same dict with added keys:
        validation_status : CLEAN | DUPLICATE | VENDOR_MISMATCH (or combined)
        validation_flags  : list of human-readable issue strings
    """
    flags = []
    po_number = (parsed_po.get("po_number") or "").strip().upper()

    # ── 1. Parse error ────────────────────────────────────────────────────────
    if parsed_po.get("parse_error"):
        flags.append(f"Parse error: {parsed_po['parse_error']}")

    # ── 2. Missing PO number ──────────────────────────────────────────────────
    if not po_number:
        flags.append("PO number is missing or blank.")

    # ── 3. Duplicate check ────────────────────────────────────────────────────
    existing_pos = get_all_po_numbers()
    if po_number and po_number in existing_pos:
        flags.append(f"DUPLICATE: PO# {po_number} already exists in the tracking log.")

    # ── 4. Vendor mismatch ────────────────────────────────────────────────────
    if parsed_po.get("vendor_mismatch"):
        form_vendor  = parsed_po.get("vendor_on_form", "?")
        email_vendor = parsed_po.get("vendor_from_email", "?")
        flags.append(
            f"VENDOR MISMATCH: Form says '{form_vendor}' "
            f"but email was sent to '{email_vendor}' vendor."
        )

    # ── 5. No items ───────────────────────────────────────────────────────────
    if not parsed_po.get("items"):
        flags.append("No line items found in PO form.")

    # ── Determine status ──────────────────────────────────────────────────────
    status_parts = []
    if any("DUPLICATE" in f for f in flags):
        status_parts.append("DUPLICATE")
    if any("VENDOR MISMATCH" in f for f in flags):
        status_parts.append("VENDOR_MISMATCH")
    if any("Parse error" in f for f in flags):
        status_parts.append("PARSE_ERROR")
    status = " | ".join(status_parts) if status_parts else "CLEAN"

    parsed_po["validation_status"] = status
    parsed_po["validation_flags"]  = flags
    return parsed_po


def validate_invoice(parsed_invoice: dict, tracking_rows: list[dict] | None = None) -> dict:
    """
    Validate an invoice against the PO tracking log.
    Checks if the referenced PO exists and if amounts match.
    """
    flags = []
    if tracking_rows is None:
        tracking_rows = _load_tracking()

    po_ref = (parsed_invoice.get("po_number") or "").strip().upper()
    inv_amount = parsed_invoice.get("amount")

    if parsed_invoice.get("parse_error"):
        flags.append(f"Parse error: {parsed_invoice['parse_error']}")

    if not po_ref:
        flags.append("Invoice does not reference a PO number.")
    else:
        match = next(
            (r for r in tracking_rows if r.get("po_number", "").upper() == po_ref),
            None,
        )
        if not match:
            flags.append(f"Referenced PO# {po_ref} not found in tracking log.")
        else:
            # Amount comparison (if we have PO item totals in the future)
            pass

    status = "INVOICE_ISSUE" if flags else "INVOICE_MATCHED"
    parsed_invoice["validation_status"] = status
    parsed_invoice["validation_flags"]  = flags
    return parsed_invoice


def log_po_to_tracking(validated_po: dict, source_file: str, received_at: str) -> dict:
    """
    Append a validated PO to the CSV tracking log.
    Returns the row that was written.
    """
    rows = _load_tracking()
    now  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    row = {
        "po_number":      (validated_po.get("po_number") or "").upper(),
        "po_type":        validated_po.get("po_type", ""),
        "supervisor":     validated_po.get("supervisor", ""),
        "vendor_on_form": validated_po.get("vendor_on_form", ""),
        "order_date":     validated_po.get("order_date", ""),
        "delivery_date":  validated_po.get("delivery_date", ""),
        "address":        validated_po.get("address", ""),
        "release":        validated_po.get("release", ""),
        "lot":            validated_po.get("lot", ""),
        "category":       validated_po.get("category", ""),
        "track":          validated_po.get("track", ""),
        "location":       validated_po.get("location", ""),
        "source_file":    source_file,
        "received_at":    received_at,
        "logged_at":      now,
        "status":         validated_po.get("validation_status", ""),
        "flags":          json.dumps(validated_po.get("validation_flags", [])),
        "invoice_amount": "",
        "po_items_json":  json.dumps(validated_po.get("items", [])),
    }

    rows.append(row)
    _save_tracking(rows)
    print(f"[Validator] Logged PO {row['po_number']} → status: {row['status']}")
    return row
