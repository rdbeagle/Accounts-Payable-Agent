"""
Agent 4: Filer & Notifier
- Saves attachment files into organized local folder structure
- Sends alert emails via Outlook COM (orders@logisticconsultants.com)
  No SMTP credentials needed — uses the Outlook account already open.
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

# ── Supervisor contact map ────────────────────────────────────────────────────
# Alerts go TO the supervisor who wrote the PO, CC to Donna.
# Update these addresses if they ever change.
SUPERVISOR_EMAILS = {
    "BK": os.getenv("EMAIL_BLAKE", "blakek@integrityllctuc.com"),
    "AN": os.getenv("EMAIL_ADAM",  "adamn@integrityllctuc.com"),
}
DONNA_EMAIL = os.getenv("EMAIL_DONNA", "donnam@logisticconsultants.com")


def _month_folder() -> str:
    return datetime.now().strftime("%Y%m")


def file_po(validated_po: dict) -> str:
    """Move/copy the PO file into the appropriate local folder."""
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


def send_alert(subject: str, body: str, to: str | None = None, cc: str | None = None) -> bool:
    """
    Send an alert email via Outlook COM.
    Sends from orders@logisticconsultants.com — the account already open
    in Outlook Classic. No SMTP credentials or App Passwords needed.
    Returns True on success, False on failure.
    """
    recipient = to or ALERT_EMAIL
    if not recipient:
        print(f"[Notifier] Alert not sent (no recipient configured): {subject}")
        return False

    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")

        mail          = outlook.CreateItem(0)   # 0 = olMailItem
        mail.Subject  = subject
        mail.Body     = body
        mail.HTMLBody = f"""<html><body>
<p style="font-family:Arial;font-size:14px;">
<strong>PO Automation Alert</strong><br><br>
{body.replace(chr(10), '<br>')}
</p>
</body></html>"""
        mail.To = recipient
        if cc:
            mail.CC = cc

        # Send from the orders account specifically
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
    Send an alert email if the document has any flags.
    Routes TO the supervisor who wrote the PO (BK=Blake, AN=Adam).
    Donna is always CC'd.
    Falls back to ALERT_EMAIL if supervisor can't be determined.
    """
    flags  = validated_doc.get("validation_flags", [])
    status = validated_doc.get("validation_status", "CLEAN")

    if not flags or status == "CLEAN":
        return

    doc_type   = validated_doc.get("file_type", "Document")
    po_number  = validated_doc.get("po_number", "N/A")
    supervisor = (validated_doc.get("supervisor") or "").upper()
    sender     = email_meta.get("sender", "Unknown")
    subject_   = email_meta.get("subject", "")
    filename   = Path(validated_doc.get("filepath", "")).name

    # Determine recipient based on supervisor code
    to_address = SUPERVISOR_EMAILS.get(supervisor, ALERT_EMAIL)
    cc_address = DONNA_EMAIL

    flag_lines = "\n".join(f"  • {f}" for f in flags)

    subject = f"[PO Alert] {status} — {po_number}"
    body = f"""A {doc_type} was received that requires your attention.

PO Number  : {po_number}
Supervisor : {supervisor if supervisor else 'Unknown'}
File       : {filename}
Received   : {email_meta.get('received_at', '')}
From       : {sender}
Subject    : {subject_}

Issues detected:
{flag_lines}

Please review the file in the 'flagged' folder and correct before resubmitting.
"""
    send_alert(subject, body, to=to_address, cc=cc_address)
