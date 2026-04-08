"""
Agent 2: Document Parser
Reads XLS/XLSX purchase order forms using the exact cell locations
established by the existing VBA macros (CompilePurchaseOrdersFromFile).

For PDFs (invoices), uses Claude to extract structured data.
"""

import os
import re
import json
import anthropic
import xlrd
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ── Vendor domain map ─────────────────────────────────────────────────────────
# Maps known vendor name keywords (from the PO form) to their email domains.
# Extend this as you onboard more vendors.
VENDOR_DOMAIN_MAP = {
    "rwc":    "rwcsupply.com",
    "l&w":    "lwsupply.com",
    "lw":     "lwsupply.com",
    "boral":  "boral.com",
    "cemex":  "cemex.com",
}


def _clean_lot(raw: str) -> str:
    """Mirrors the VBA CleanLot function."""
    temp = re.sub(r"(?i)lot|#", "", raw)
    return re.sub(r"[^A-Za-z0-9]", "", temp).strip()


def _detect_supervisor(po_number: str) -> str:
    """Last 2 chars of PO# = supervisor code (BK or AN)."""
    if len(po_number) >= 2:
        suffix = po_number[-2:].upper()
        return suffix if suffix in ("BK", "AN") else "?"
    return "?"


def _vendor_from_form(ws, po_type: str) -> str:
    """
    Extract vendor name from the PO form.
    Lath forms: check region B7:I10 for vendor references.
    Stucco forms: check B7:I12.
    This reads cell B7 which typically contains vendor/track info.
    """
    try:
        return str(ws.cell_value(6, 1)).strip()  # B7 (0-indexed: row 6, col 1)
    except Exception:
        return ""


def parse_xls_po(filepath: str, to_addresses: list[str]) -> dict:
    """
    Parse a Lath or Stucco PO from an XLS file.
    Returns a structured dict matching the compiler's output schema.
    """
    result = {
        "filepath":      filepath,
        "file_type":     "PO",
        "parse_error":   None,
        "po_type":       None,
        "po_number":     None,
        "supervisor":    None,
        "order_date":    None,
        "delivery_date": None,
        "address":       None,
        "release":       None,
        "lot":           None,
        "category":      None,
        "track":         None,
        "location":      None,
        "vendor_on_form": None,
        "vendor_from_email": None,
        "vendor_mismatch": False,
        "items":         [],
    }

    try:
        wb = xlrd.open_workbook(filepath)
        # The compiler always reads the "Form" sheet
        try:
            ws = wb.sheet_by_name("Form")
        except xlrd.biffh.XLRDError:
            ws = wb.sheets()[0]

        # ── Detect PO type from A12 (0-indexed: row 11, col 0) ──────────────
        a12 = str(ws.cell_value(11, 0)).lower().strip()
        if "lath" in a12:
            po_type = "Lath"
            po_number = str(ws.cell_value(11, 8)).strip().upper()   # I12
            order_date    = str(ws.cell_value(5, 8)).strip()         # I6
            delivery_date = str(ws.cell_value(6, 8)).strip()         # I7
            address       = str(ws.cell_value(9, 7)).strip()         # H10
            category      = "MT10"
            item_start, item_end = 13, 46   # rows 14–46 (0-indexed)
            track         = str(ws.cell_value(6, 1)).strip()         # B7
        elif "stucco" in a12:
            po_type = "Stucco"
            po_number = str(ws.cell_value(11, 5)).strip().upper()   # F12
            order_date    = str(ws.cell_value(5, 7)).strip()         # H6
            delivery_date = str(ws.cell_value(6, 7)).strip()         # H7
            address       = str(ws.cell_value(10, 6)).strip()        # G11
            category      = "MT40"
            item_start, item_end = 13, 29   # rows 14–29 (0-indexed)
            track         = str(ws.cell_value(6, 1)).strip()         # B7
        else:
            result["parse_error"] = f"Unknown PO type in A12: '{a12}'"
            return result

        # ── WILLCALL handling ────────────────────────────────────────────────
        if order_date.upper() == "WILLCALL":
            order_date = "WILLCALL"
        if delivery_date.upper() in ("WILLCALL", ""):
            delivery_date = order_date

        # ── Lot ──────────────────────────────────────────────────────────────
        raw_lot = str(ws.cell_value(8, 3)).strip()   # D9
        lot = _clean_lot(raw_lot)

        # ── Release ──────────────────────────────────────────────────────────
        release = str(ws.cell_value(7, 1)).strip()   # B8

        # ── Location (job vs yard) ────────────────────────────────────────────
        # Scan the header region for the word "yard"
        location = "JOBTUC"
        try:
            end_col = 9 if po_type == "Lath" else 9
            end_row = 10 if po_type == "Lath" else 12
            for r in range(6, end_row):
                for c in range(1, end_col):
                    if "yard" in str(ws.cell_value(r, c)).lower():
                        location = "TUCSON"
        except Exception:
            pass

        # ── Items ─────────────────────────────────────────────────────────────
        items = []
        for r in range(item_start, min(item_end, ws.nrows)):
            desc = str(ws.cell_value(r, 0)).strip()
            try:
                qty = float(ws.cell_value(r, 1))
            except (ValueError, TypeError):
                qty = 0
            if desc and qty > 0:
                items.append({"description": desc, "quantity": qty})

        # ── Vendor extraction ─────────────────────────────────────────────────
        vendor_on_form = _vendor_from_form(ws, po_type)

        # Determine vendor implied by the email's To: addresses
        vendor_from_email = None
        for addr in to_addresses:
            domain = addr.split("@")[-1].lower() if "@" in addr else ""
            for keyword, vdomain in VENDOR_DOMAIN_MAP.items():
                if vdomain in domain or domain in vdomain:
                    vendor_from_email = keyword.upper()
                    break
            if vendor_from_email:
                break

        # ── Mismatch detection ────────────────────────────────────────────────
        vendor_mismatch = False
        if vendor_on_form and vendor_from_email:
            form_lower  = vendor_on_form.lower()
            email_lower = vendor_from_email.lower()
            # Flag if the vendor named on the form doesn't match who the email was sent to
            vendor_mismatch = not any(
                kw in form_lower for kw in email_lower.split()
            ) and not any(
                kw in email_lower for kw in form_lower.split()
            )

        result.update({
            "po_type":           po_type,
            "po_number":         po_number,
            "supervisor":        _detect_supervisor(po_number),
            "order_date":        order_date,
            "delivery_date":     delivery_date,
            "address":           address,
            "release":           release,
            "lot":               lot,
            "category":          category,
            "track":             track,
            "location":          location,
            "vendor_on_form":    vendor_on_form,
            "vendor_from_email": vendor_from_email,
            "vendor_mismatch":   vendor_mismatch,
            "items":             items,
        })

    except Exception as e:
        result["parse_error"] = str(e)

    return result


def parse_pdf_invoice(filepath: str, email_metadata: dict) -> dict:
    """
    Use Claude to extract structured invoice data from a PDF.
    Returns a dict with invoice_number, vendor, po_number, amount, line_items.
    """
    result = {
        "filepath":       filepath,
        "file_type":      "Invoice",
        "parse_error":    None,
        "invoice_number": None,
        "vendor":         None,
        "po_number":      None,
        "amount":         None,
        "line_items":     [],
    }

    try:
        import pdfplumber
        with pdfplumber.open(filepath) as pdf:
            text = "\n".join(
                page.extract_text() or "" for page in pdf.pages
            )
    except Exception as e:
        result["parse_error"] = f"PDF read failed: {e}"
        return result

    if not text.strip():
        result["parse_error"] = "PDF contained no extractable text."
        return result

    try:
        client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
        prompt = f"""You are a document parser for a construction supply company.
Extract the following fields from this invoice text and return ONLY valid JSON.

Required fields:
- invoice_number (string)
- vendor (string — the company issuing the invoice)
- po_number (string — the purchase order this invoice references, if present)
- amount (number — total invoice amount)
- line_items (array of objects with: description, quantity, unit_price)

Invoice text:
{text[:6000]}

Return only JSON. No explanation, no markdown fences."""

        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1000,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = response.content[0].text.strip()
        raw = re.sub(r"^```json|^```|```$", "", raw, flags=re.MULTILINE).strip()
        parsed = json.loads(raw)
        result.update(parsed)
        result["file_type"] = "Invoice"
        result["filepath"]  = filepath

    except Exception as e:
        result["parse_error"] = f"Claude extraction failed: {e}"

    return result
