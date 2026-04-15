"""
Agent 3: Validator
- Checks for duplicate PO numbers
- Flags vendor mismatches
- Matches invoices to POs using Claude for extraction + rule-based checking
"""

import os
import csv
import json
import re
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
    "status", "flags", "invoice_amount", "po_items_json",
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
    flags     = []
    po_number = (parsed_po.get("po_number") or "").strip().upper()

    if parsed_po.get("parse_error"):
        flags.append(f"Parse error: {parsed_po['parse_error']}")

    if not po_number:
        flags.append("PO number is missing or blank.")

    existing_pos = get_all_po_numbers()
    if po_number and po_number in existing_pos:
        flags.append(f"DUPLICATE: PO# {po_number} already exists in the tracking log.")

    if parsed_po.get("vendor_mismatch"):
        form_vendor  = parsed_po.get("vendor_on_form", "?")
        email_vendor = parsed_po.get("vendor_from_email", "?")
        flags.append(
            f"VENDOR MISMATCH: Form says '{form_vendor}' "
            f"but email was sent to '{email_vendor}' vendor."
        )

    if not parsed_po.get("items"):
        flags.append("No line items found in PO form.")

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


def _extract_invoice_fields_with_claude(filepath: str) -> dict:
    """
    Use Claude to extract vendor name and referenced PO number from an invoice PDF.
    Returns dict with: vendor, po_number_reference, invoice_number
    """
    result = {"vendor": None, "po_number_reference": None, "invoice_number": None}
    try:
        import pdfplumber
        import anthropic

        with pdfplumber.open(filepath) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)

        if not text.strip():
            return result

        client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
        prompt = f"""Extract these fields from this invoice and return ONLY valid JSON:
- invoice_number (string)
- vendor (string — company issuing the invoice)
- po_number_reference (string — the customer's PO number referenced on this invoice, often labeled "Customer PO#", "PO Number", "Your PO", or similar. This is NOT the invoice number.)

Invoice text:
{text[:4000]}

Return only JSON. No explanation."""

        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=300,
            messages=[{"role": "user", "content": prompt}]
        )
        raw = response.content[0].text.strip()
        raw = re.sub(r"^```json|^```|```$", "", raw, flags=re.MULTILINE).strip()
        parsed = json.loads(raw)
        result.update(parsed)

    except Exception as e:
        print(f"[Validator] Invoice Claude extraction failed: {e}")

    return result


def validate_invoice(parsed_invoice: dict, tracking_rows: list[dict] | None = None) -> dict:
    """
    Validate an invoice:
    1. Use Claude to extract vendor and PO reference from the PDF
    2. Check if that PO exists in the tracking log
    3. Check if the vendor matches
    """
    flags = []
    if tracking_rows is None:
        tracking_rows = _load_tracking()

    if parsed_invoice.get("parse_error"):
        flags.append(f"Parse error: {parsed_invoice['parse_error']}")
        parsed_invoice["validation_status"] = "INVOICE_ISSUE"
        parsed_invoice["validation_flags"]  = flags
        return parsed_invoice

    filepath = parsed_invoice.get("filepath", "")

    # Use Claude to extract fields from the invoice PDF
    extracted = _extract_invoice_fields_with_claude(filepath)
    po_ref    = (extracted.get("po_number_reference") or "").strip().upper()
    inv_vendor = (extracted.get("vendor") or "").strip()
    inv_number = extracted.get("invoice_number", "")

    # Store extracted fields
    parsed_invoice["po_number"]      = po_ref
    parsed_invoice["vendor"]         = inv_vendor
    parsed_invoice["invoice_number"] = inv_number

    if not po_ref:
        flags.append(f"Invoice from {inv_vendor or 'unknown vendor'} does not reference a PO number.")
    else:
        # Check if PO exists in tracking
        match = next(
            (r for r in tracking_rows if r.get("po_number", "").upper() == po_ref),
            None
        )
        if not match:
            flags.append(
                f"Invoice references PO# {po_ref} but this PO is not in the tracking log. "
                f"Either the PO was not submitted or has not been processed yet."
            )
        else:
            # Vendor cross-check
            po_vendor = (match.get("vendor_on_form") or "").strip().lower()
            if po_vendor and inv_vendor:
                inv_vendor_lower = inv_vendor.lower()
                if not any(
                    kw in inv_vendor_lower for kw in po_vendor.split()
                ) and not any(
                    kw in po_vendor for kw in inv_vendor_lower.split()
                ):
                    flags.append(
                        f"Vendor on invoice ('{inv_vendor}') does not match "
                        f"vendor on PO# {po_ref} ('{match.get('vendor_on_form')}')."
                    )
            else:
                print(f"[Validator] Invoice PO# {po_ref} matched → {match.get('po_number')}")

    status = "INVOICE_ISSUE" if flags else "INVOICE_MATCHED"
    parsed_invoice["validation_status"] = status
    parsed_invoice["validation_flags"]  = flags
    return parsed_invoice


def log_po_to_tracking(validated_po: dict, source_file: str, received_at: str) -> dict:
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
