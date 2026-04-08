"""
Agent 4: Filer & Notifier
- Saves attachment files into organized local folder structure
- Sends alert emails for duplicates, vendor mismatches, or invoice issues
"""

import os
import shutil
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

STORAGE_ROOT  = os.getenv("STORAGE_ROOT", "./data")
ALERT_EMAIL   = os.getenv("ALERT_EMAIL", "")
SMTP_HOST     = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT     = int(os.getenv("SMTP_PORT", 587))
SMTP_USER     = os.getenv("SMTP_USER", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")


# ── Folder layout ─────────────────────────────────────────────────────────────
# data/
#   processed/
#     YYYYMM/
#       POs/
#         {po_number}_{filename}
#       Invoices/
#         {filename}
#   flagged/
#     {po_number}_{filename}
#   archive/
#     inbox copies (untouched originals)

def _month_folder() -> str:
    return datetime.now().strftime("%Y%m")


def file_po(validated_po: dict) -> str:
    """
    Move/copy the PO file into the appropriate local folder.
    Returns the destination path.
    """
    src = validated_po.get("filepath", "")
    if not src or not Path(src).exists():
        print(f"[Filer] Source file not found: {src}")
        return ""

    po_number = (validated_po.get("po_number") or "UNKNOWN").upper()
    status    = validated_po.get("validation_status", "")
    filename  = Path(src).name

    if status and status != "CLEAN":
        dest_dir = Path(STORAGE_ROOT) / "flagged"
    else:
        dest_dir = Path(STORAGE_ROOT) / "processed" / _month_folder() / "POs"

    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / f"{po_number}_{filename}"

    # Avoid overwriting: append counter if file already exists
    counter = 1
    while dest.exists():
        stem = f"{po_number}_{Path(filename).stem}_{counter}"
        dest = dest_dir / f"{stem}{Path(filename).suffix}"
        counter += 1

    shutil.copy2(src, dest)
    print(f"[Filer] PO filed → {dest}")
    return str(dest)


def file_invoice(validated_invoice: dict) -> str:
    """Move/copy the invoice PDF into the appropriate folder."""
    src = validated_invoice.get("filepath", "")
    if not src or not Path(src).exists():
        return ""

    status   = validated_invoice.get("validation_status", "")
    filename = Path(src).name

    if status and status != "INVOICE_MATCHED":
        dest_dir = Path(STORAGE_ROOT) / "flagged"
    else:
        dest_dir = Path(STORAGE_ROOT) / "processed" / _month_folder() / "Invoices"

    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / filename

    counter = 1
    while dest.exists():
        stem = f"{Path(filename).stem}_{counter}"
        dest = dest_dir / f"{stem}{Path(filename).suffix}"
        counter += 1

    shutil.copy2(src, dest)
    print(f"[Filer] Invoice filed → {dest}")
    return str(dest)


def send_alert(subject: str, body: str, to: str | None = None) -> bool:
    """
    Send an alert email via SMTP.
    Returns True on success, False on failure.
    """
    recipient = to or ALERT_EMAIL
    if not recipient or not SMTP_USER or not SMTP_PASSWORD:
        print(f"[Notifier] Alert not sent (SMTP not configured): {subject}")
        return False

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = SMTP_USER
        msg["To"]      = recipient

        html_body = f"""
<html><body>
<p style="font-family:Arial;font-size:14px;">
<strong>PO Automation Alert</strong><br><br>
{body.replace(chr(10), '<br>')}
</p>
</body></html>"""

        msg.attach(MIMEText(body, "plain"))
        msg.attach(MIMEText(html_body, "html"))

        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, recipient, msg.as_string())

        print(f"[Notifier] Alert sent to {recipient}: {subject}")
        return True

    except Exception as e:
        print(f"[Notifier] Failed to send alert: {e}")
        return False


def notify_if_flagged(validated_doc: dict, email_meta: dict):
    """
    Send an alert email if the document has any flags.
    """
    flags  = validated_doc.get("validation_flags", [])
    status = validated_doc.get("validation_status", "CLEAN")

    if not flags or status == "CLEAN":
        return

    doc_type  = validated_doc.get("file_type", "Document")
    po_number = validated_doc.get("po_number", "N/A")
    sender    = email_meta.get("sender", "Unknown")
    subject_  = email_meta.get("subject", "")
    filename  = Path(validated_doc.get("filepath", "")).name

    flag_lines = "\n".join(f"  • {f}" for f in flags)

    subject = f"[PO Alert] {status} — {po_number} from {sender}"
    body = f"""A {doc_type} was received that requires attention.

PO Number : {po_number}
File      : {filename}
Received  : {email_meta.get('received_at', '')}
From      : {sender}
Subject   : {subject_}

Issues detected:
{flag_lines}

Please review the file in the 'flagged' folder.
"""
    send_alert(subject, body)
