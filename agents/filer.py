"""
Agent 4: Filer & Notifier
- Saves attachment files into organized local folder structure
- Sends alert emails via Outlook COM (orders@logisticconsultants.com)
- Uses Claude to write contextual, actionable alert messages
"""

import os
import shutil
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

STORAGE_ROOT  = os.getenv("STORAGE_ROOT", "./data")
ALERT_EMAIL   = os.getenv("ALERT_EMAIL", "orders@logisticconsultants.com")
OUTLOOK_INBOX = os.getenv("OUTLOOK_INBOX", "orders@logisticconsultants.com")

SUPERVISOR_EMAILS = {
    "BK": os.getenv("EMAIL_BLAKE", "blakek@integrityllctuc.com"),
    "AN": os.getenv("EMAIL_ADAM",  "adamn@integrityllctuc.com"),
}
DONNA_EMAIL = os.getenv("EMAIL_DONNA", "donnam@logisticconsultants.com")


def _month_folder() -> str:
    return datetime.now().strftime("%Y%m")


def file_po(validated_po: dict) -> str:
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

    counter = 1
    while dest.exists():
        stem = f"{po_number}_{Path(filename).stem}_{counter}"
        dest = dest_dir / f"{stem}{Path(filename).suffix}"
        counter += 1

    shutil.copy2(src, dest)
    print(f"[Filer] PO filed → {dest}")
    return str(dest)


def file_invoice(validated_invoice: dict) -> str:
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


def _generate_alert_body(validated_doc: dict, email_meta: dict, flags: list) -> str:
    """
    Use Claude to write a contextual, actionable alert email body.
    Falls back to a plain template if the API call fails.
    """
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

        doc_type   = validated_doc.get("file_type", "Document")
        po_number  = validated_doc.get("po_number", "N/A")
        supervisor = (validated_doc.get("supervisor") or "Unknown").upper()
        status     = validated_doc.get("validation_status", "")
        vendor     = validated_doc.get("vendor_on_form", "Unknown")
        email_vendor = validated_doc.get("vendor_from_email", "Unknown")
        filename   = Path(validated_doc.get("filepath", "")).name
        sender     = email_meta.get("sender", "Unknown")
        subject    = email_meta.get("subject", "")
        flag_text  = "\n".join(f"- {f}" for f in flags)

        prompt = f"""You are writing a brief alert email for a construction supply company's purchase order system.

A {doc_type} was received that has issues requiring attention. Write a 3-4 sentence plain English message that:
1. Clearly states what the problem is
2. Explains what likely caused it
3. States exactly what action needs to be taken to fix it

Be direct and professional. Do not use bullet points. Do not include a subject line or greeting.

Details:
- PO Number: {po_number}
- Supervisor Code: {supervisor}
- Status: {status}
- Vendor on Form: {vendor}
- Email sent to vendor: {email_vendor}
- File: {filename}
- Email from: {sender}
- Email subject: {subject}
- Issues detected:
{flag_text}

Write only the email body, 3-4 sentences."""

        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=300,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text.strip()

    except Exception as e:
        print(f"[Notifier] Claude alert generation failed, using template: {e}")
        # Fall back to plain template
        flag_lines = "\n".join(f"  • {f}" for f in flags)
        return f"""A {validated_doc.get('file_type','Document')} was received that requires attention.

PO Number  : {validated_doc.get('po_number', 'N/A')}
File       : {Path(validated_doc.get('filepath','')).name}
From       : {email_meta.get('sender','')}

Issues detected:
{flag_lines}

Please review the file in the 'flagged' folder and correct before resubmitting."""


def send_alert(subject: str, body: str, to: str | None = None, cc: str | None = None) -> bool:
    """
    Send an alert email via Outlook COM.
    Sends from orders@logisticconsultants.com — no SMTP credentials needed.
    """
    recipient = to or ALERT_EMAIL
    if not recipient:
        print(f"[Notifier] Alert not sent (no recipient configured): {subject}")
        return False

    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")

        mail          = outlook.CreateItem(0)
        mail.Subject  = subject
        mail.Body     = body
        mail.HTMLBody = f"""<html><body>
<p style="font-family:Arial;font-size:14px;">
<strong>PO Automation Alert — Integrity Wall Systems</strong><br><br>
{body.replace(chr(10), '<br>')}
</p>
</body></html>"""
        mail.To = recipient
        if cc:
            mail.CC = cc

        ns = outlook.GetNamespace("MAPI")
        for account in ns.Accounts:
            try:
                if OUTLOOK_INBOX.lower() in account.SmtpAddress.lower():
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                    break
            except Exception:
                continue

        mail.Send()
        cc_note = f", CC: {cc}" if cc else ""
        print(f"[Notifier] Alert sent to {recipient}{cc_note}: {subject}")
        return True

    except Exception as e:
        print(f"[Notifier] Failed to send alert via Outlook: {e}")
        return False


def notify_if_flagged(validated_doc: dict, email_meta: dict):
    """
    Send a Claude-written alert email if the document has any flags.
    Routes TO the supervisor (BK=Blake, AN=Adam), CC Donna.
    """
    flags  = validated_doc.get("validation_flags", [])
    status = validated_doc.get("validation_status", "CLEAN")

    if not flags or status == "CLEAN":
        return

    po_number  = validated_doc.get("po_number", "N/A")
    supervisor = (validated_doc.get("supervisor") or "").upper()

    to_address = SUPERVISOR_EMAILS.get(supervisor, ALERT_EMAIL)
    cc_address = DONNA_EMAIL

    # Generate contextual body via Claude
    body = _generate_alert_body(validated_doc, email_meta, flags)

    subject = f"[PO Alert] {status} — {po_number}"
    send_alert(subject, body, to=to_address, cc=cc_address)
