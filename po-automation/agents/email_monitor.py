"""
Agent 1: Email Monitor
Connects to the shared inbox via IMAP, fetches unread emails that have
attachments, saves attachments locally, and returns structured metadata.
"""

import imaplib
import email
import os
import re
from email.header import decode_header
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

IMAP_HOST     = os.getenv("IMAP_HOST", "mail.logisticconsultants.com")
IMAP_PORT     = int(os.getenv("IMAP_PORT", 993))
IMAP_USER     = os.getenv("IMAP_USER", "")
IMAP_PASSWORD = os.getenv("IMAP_PASSWORD", "")
IMAP_MAILBOX  = os.getenv("IMAP_MAILBOX", "INBOX")
STORAGE_ROOT  = os.getenv("STORAGE_ROOT", "./data")

SUPPORTED_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".pdf"}


def _decode_str(value: str | bytes) -> str:
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return value or ""


def _safe_filename(name: str) -> str:
    """Strip characters that are unsafe in file names."""
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    return name.strip()


def _extract_to_addresses(msg) -> list[str]:
    """Return all recipient email addresses from To/CC headers."""
    addresses = []
    for header in ("To", "CC"):
        raw = msg.get(header, "")
        if raw:
            # Pull bare addresses out of "Name <addr>" or plain "addr" formats
            addresses += re.findall(r"[\w.\-+]+@[\w.\-]+", raw)
    return [a.lower() for a in addresses]


def fetch_new_emails(mark_seen: bool = True) -> list[dict]:
    """
    Connect to IMAP, fetch UNSEEN emails, and return a list of dicts:
    {
        uid, subject, sender, to_addresses, received_at,
        attachments: [{ filename, filepath, extension }]
    }
    Only emails with at least one supported attachment are returned.
    """
    results = []
    inbox_path = Path(STORAGE_ROOT) / "inbox"
    inbox_path.mkdir(parents=True, exist_ok=True)

    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(IMAP_USER, IMAP_PASSWORD)
        mail.select(IMAP_MAILBOX)
    except Exception as e:
        print(f"[EmailMonitor] Connection failed: {e}")
        return []

    try:
        status, data = mail.search(None, "UNSEEN")
        if status != "OK" or not data[0]:
            print("[EmailMonitor] No unseen emails.")
            mail.logout()
            return []

        uid_list = data[0].split()
        print(f"[EmailMonitor] Found {len(uid_list)} unseen email(s).")

        for uid in uid_list:
            status, msg_data = mail.fetch(uid, "(RFC822)")
            if status != "OK":
                continue

            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            # Decode subject
            subject_parts = decode_header(msg.get("Subject", ""))
            subject = " ".join(
                _decode_str(part) if isinstance(part, bytes) else (part or "")
                for part, enc in subject_parts
            )

            sender      = msg.get("From", "")
            to_addresses = _extract_to_addresses(msg)
            date_str    = msg.get("Date", "")

            attachments = []
            for part in msg.walk():
                content_disposition = part.get("Content-Disposition", "")
                if "attachment" not in content_disposition.lower():
                    continue

                filename = part.get_filename()
                if not filename:
                    continue

                # Decode RFC 2047 encoded filenames
                decoded_parts = decode_header(filename)
                filename = " ".join(
                    _decode_str(p) if isinstance(p, bytes) else (p or "")
                    for p, enc in decoded_parts
                )
                filename = _safe_filename(filename)
                ext = Path(filename).suffix.lower()

                if ext not in SUPPORTED_EXTENSIONS:
                    print(f"[EmailMonitor] Skipping unsupported file: {filename}")
                    continue

                # Save attachment with timestamp prefix to avoid collisions
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                save_name = f"{ts}_{filename}"
                save_path = inbox_path / save_name
                save_path.write_bytes(part.get_payload(decode=True))

                attachments.append({
                    "filename":  filename,
                    "filepath":  str(save_path),
                    "extension": ext,
                })
                print(f"[EmailMonitor] Saved attachment: {save_name}")

            if not attachments:
                continue  # skip emails with no supported attachments

            results.append({
                "uid":          uid.decode(),
                "subject":      subject,
                "sender":       sender,
                "to_addresses": to_addresses,
                "received_at":  date_str,
                "attachments":  attachments,
            })

            if mark_seen:
                mail.store(uid, "+FLAGS", "\\Seen")

    finally:
        mail.logout()

    print(f"[EmailMonitor] Returning {len(results)} email(s) with attachments.")
    return results
