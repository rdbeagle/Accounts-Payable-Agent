"""
Agent 2: Document Parser
Reads XLS/XLSX purchase order forms using the exact cell locations
established by the existing VBA macros (CompilePurchaseOrdersFromFile).

Cell reference map (confirmed from real form inspection):
  Both form types:
    A12  = form type ("lath" or "stucco")
    B7   = tract/neighborhood (Arbor at Madera, Quail Creek, etc.) — NOT vendor
    B8   = release number
    D9   = lot number
    G10  = vendor code (LW, L&W, DBM, RWC, etc.)  ← key fix
  Lath:
    I12  = PO number
    I6   = order date, I7 = delivery date
    H10  = address
    rows 14-46 = line items (col A=desc, col B=qty)
  Stucco:
    F12  = PO number
    H6   = order date, H7 = delivery date
    G11  = address
    rows 14-29 = line items

For PDFs (invoices), uses Claude AI to extract structured data.
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
# Maps vendor codes written in G10 on the PO form to their email domains.
# LW, L&W, and DBM all refer to L&W Supply Corporation.
# Add new vendors here as you confirm their codes and domains.
VENDOR_DOMAIN_MAP = {
    "rwc":   "rwcsupply.com",
    "l&w":   "lwsupply.com",
    "lw":    "lwsupply.com",
    "dbm":   "lwsupply.com",   # DBM is an alias for L&W Supply
    "boral": "boral.com",
    "cemex": "cemex.com",
}

# Reverse map: email domain → canonical vendor name for display
DOMAIN_TO_VENDOR = {
    "rwcsupply.com":  "RWC",
    "rwc.org":        "RWC",
    "lwsupply.com":   "L&W",
    "boral.com":      "Boral",
    "cemex.com":      "Cemex",
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


def _vendor_from_form(ws) -> str:
    """
    Find the vendor value by searching for the label cell containing
    the word 'vendor' and reading the cell immediately to its right.
    This handles both Lath and Stucco form layouts which differ by row.

    Stucco: F10='VENDOR', G10=value
    Lath:   F11='Vendor', G11=value
    Searches rows 8-14, columns 0-8 to find the label.
    """
    try:
        for r in range(8, 14):
            for c in range(0, 9):
                cell = str(ws.cell_value(r, c)).strip().lower()
                if cell == "vendor":
                    # Return the cell immediately to the right
                    value = str(ws.cell_value(r, c + 1)).strip()
                    if value:
                        return value
    except Exception:
        pass
    return ""


def _vendor_from_email(to_addresses: list[str]) -> str | None:
    """
    Determine vendor implied by who the email was sent to.
    Checks To: addresses against known vendor domains.
    Returns canonical vendor name or None.
    """
    for addr in to_addresses:
        domain = addr.split("@")[-1].lower() if "@" in addr else ""
        # Check direct domain match
        for vdomain, vname in DOMAIN_TO_VENDOR.items():
            if vdomain in domain or domain in vdomain:
                return vname
    return None


def parse_xls_po(filepath: str, to_addresses: list[str]) -> dict:
    """
    Parse a Lath or Stucco PO from an XLS file.
    Returns a structured dict matching the compiler's output schema.
    """
    result = {
        "filepath":          filepath,
        "file_type":         "PO",
        "parse_error":       None,
        "po_type":           None,
        "po_number":         None,
        "supervisor":        None,
        "order_date":        None,
        "delivery_date":     None,
        "address":           None,
        "release":           None,
        "lot":               None,
        "tract":             None,   # B7 — neighborhood/phase name
        "category":          None,
        "location":          None,
        "vendor_on_form":    None,   # G10 — actual vendor code
        "vendor_from_email": None,   # inferred from To: domain
        "vendor_mismatch":   False,
        "items":             [],
    }

    try:
        try:
            wb = xlrd.open_workbook(filepath)
        except Exception as e:
            if "xlsx" in str(e).lower() or "not supported" in str(e).lower():
                result["parse_error"] = "Not a PO form (xlsx format — inventory/report file)"
            else:
                result["parse_error"] = str(e)
            return result
        try:
            ws = wb.sheet_by_name("Form")
        except xlrd.biffh.XLRDError:
            ws = wb.sheets()[0]

        # ── Detect PO type from A12 ──────────────────────────────────────────
        a12 = str(ws.cell_value(11, 0)).lower().strip() if ws.nrows > 11 else ""
        if "lath" in a12:
            po_type       = "Lath"
            po_number     = str(ws.cell_value(11, 8)).strip().upper()  # I12
            order_date    = str(ws.cell_value(5, 8)).strip()           # I6
            delivery_date = str(ws.cell_value(6, 8)).strip()           # I7
            address       = str(ws.cell_value(9, 7)).strip()           # H10
            category      = "MT10"
            item_start, item_end = 13, 46
        elif "stucco" in a12:
            po_type       = "Stucco"
            po_number     = str(ws.cell_value(11, 5)).strip().upper()  # F12
            order_date    = str(ws.cell_value(5, 7)).strip()           # H6
            delivery_date = str(ws.cell_value(6, 7)).strip()           # H7
            address       = str(ws.cell_value(10, 6)).strip()          # G11
            category      = "MT40"
            item_start, item_end = 13, 29
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

        # ── Tract (neighborhood/phase name) ──────────────────────────────────
        tract = str(ws.cell_value(6, 1)).strip()     # B7

        # ── Location (job vs yard) ────────────────────────────────────────────
        location = "JOBTUC"
        try:
            end_row = 10 if po_type == "Lath" else 12
            for r in range(6, end_row):
                for c in range(1, 9):
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

        # ── Vendor from G10 ───────────────────────────────────────────────────
        vendor_on_form = _vendor_from_form(ws)

        # ── Vendor implied by email To: addresses ─────────────────────────────
        email_vendor = _vendor_from_email(to_addresses)

        # ── Mismatch detection ────────────────────────────────────────────────
        # Compare G10 vendor code against who the email was actually sent to.
        vendor_mismatch = False
        if vendor_on_form and email_vendor:
            # Normalise form vendor: strip spaces, lowercase
            form_norm = vendor_on_form.lower().replace(" ", "").replace("&", "")
            # Map form code → canonical name using VENDOR_DOMAIN_MAP
            form_canonical = None
            for code, domain in VENDOR_DOMAIN_MAP.items():
                if code in form_norm or form_norm in code:
                    # Find canonical name from domain
                    form_canonical = DOMAIN_TO_VENDOR.get(domain, code.upper())
                    break

            if form_canonical and form_canonical != email_vendor:
                vendor_mismatch = True
            elif not form_canonical:
                # Unknown vendor code on form — flag for review
                vendor_mismatch = True

        result.update({
            "po_type":           po_type,
            "po_number":         po_number,
            "supervisor":        _detect_supervisor(po_number),
            "order_date":        order_date,
            "delivery_date":     delivery_date,
            "address":           address,
            "release":           release,
            "lot":               lot,
            "tract":             tract,
            "category":          category,
            "location":          location,
            "vendor_on_form":    vendor_on_form,
            "vendor_from_email": email_vendor,
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
