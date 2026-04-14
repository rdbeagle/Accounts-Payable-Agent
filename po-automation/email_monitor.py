"""
Agent 1: Email Monitor (Outlook COM version)
Reads directly from Outlook Classic on this machine — no IMAP credentials needed.
Requires: pywin32, Outlook installed and open with the orders inbox loaded.
"""

import os
import re
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

STORAGE_ROOT = os.getenv("STORAGE_ROOT", "./data")
SUPPORTED_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".pdf"}

# Name of the inbox to monitor.
# If None, uses the default inbox. Set to "orders@logisticconsultants.com"
# if you have multiple accounts in Outlook and want to target the shared one.
TARGET_INBOX = os.getenv("OUTLOOK_INBOX", None)


def _safe_filename(name: str) -> str:
    """Strip characters unsafe in file names."""
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


def _get_inbox(outlook_ns):
    """
    Return the correct inbox folder.
    If TARGET_INBOX is set, search all accounts for a matching store,
    then return its Inbox folder by name.
    Otherwise return the default inbox.
    """
    if TARGET_INBOX:
        for store in outlook_ns.Stores:
            try:
                if TARGET_INBOX.lower() in store.DisplayName.lower():
                    root = store.GetRootFolder()
                    for folder in root.Folders:
                        if folder.Name.lower() == "inbox":
                            print(f"[EmailMonitor] Using inbox: {store.DisplayName} → {folder.Name} ({folder.UnReadItemCount} unread)")
                            return folder
            except Exception:
                continue
        print(f"[EmailMonitor] Warning: Could not find inbox for '{TARGET_INBOX}', falling back to default.")

    # Default inbox (folder 6 = olFolderInbox)
    return outlook_ns.GetDefaultFolder(6)


def _extract_to_addresses(message) -> list[str]:
    """Extract all recipient email addresses from a message."""
    addresses = []
    try:
        for i in range(1, message.Recipients.Count + 1):
            recipient = message.Recipients.Item(i)
            addr = recipient.Address or ""
            # Outlook sometimes gives Exchange addresses — try to get SMTP
            try:
                smtp = recipient.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                )
                if smtp:
                    addr = smtp
            except Exception:
                pass
            if "@" in addr:
                addresses.append(addr.lower())
    except Exception as e:
        print(f"[EmailMonitor] Could not read recipients: {e}")
    return addresses


def fetch_new_emails(mark_seen: bool = True) -> list[dict]:
    """
    Connect to Outlook via COM, fetch unread emails with supported attachments.

    Returns a list of dicts:
    {
        uid, subject, sender, to_addresses, received_at,
        attachments: [{ filename, filepath, extension }]
    }
    """
    try:
        import win32com.client
    except ImportError:
        print("[EmailMonitor] ERROR: pywin32 not installed. Run: pip install pywin32")
        return []

    results = []
    inbox_path = Path(STORAGE_ROOT) / "inbox"
    inbox_path.mkdir(parents=True, exist_ok=True)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"[EmailMonitor] Could not connect to Outlook: {e}")
        print("  Make sure Outlook is open and signed in.")
        return []

    try:
        inbox = _get_inbox(ns)
    except Exception as e:
        print(f"[EmailMonitor] Could not access inbox: {e}")
        return []

    # Filter to unread messages only
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)  # newest first

    unread_count = 0
    processed_count = 0

    for message in messages:
        try:
            # Only process unread emails
            if message.UnRead is False:
                continue

            unread_count += 1

            # Skip if no attachments
            if message.Attachments.Count == 0:
                if mark_seen:
                    message.UnRead = False
                    message.Save()
                continue

            subject     = message.Subject or ""
            sender      = message.SenderEmailAddress or message.SenderName or ""
            received_at = str(message.ReceivedTime)
            uid         = str(message.EntryID)

            to_addresses = _extract_to_addresses(message)

            attachments = []
            for i in range(1, message.Attachments.Count + 1):
                attachment = message.Attachments.Item(i)
                filename   = _safe_filename(attachment.FileName or f"attachment_{i}")
                ext        = Path(filename).suffix.lower()

                if ext not in SUPPORTED_EXTENSIONS:
                    print(f"[EmailMonitor] Skipping unsupported file: {filename}")
                    continue

                # Save with timestamp prefix to avoid collisions
                ts        = datetime.now().strftime("%Y%m%d_%H%M%S")
                save_name = f"{ts}_{filename}"
                save_path = inbox_path / save_name

                attachment.SaveAsFile(str(save_path))
                print(f"[EmailMonitor] Saved: {save_name}")

                attachments.append({
                    "filename":  filename,
                    "filepath":  str(save_path),
                    "extension": ext,
                })

            if not attachments:
                # Had attachments but none were supported types
                if mark_seen:
                    message.UnRead = False
                    message.Save()
                continue

            results.append({
                "uid":          uid,
                "subject":      subject,
                "sender":       sender,
                "to_addresses": to_addresses,
                "received_at":  received_at,
                "attachments":  attachments,
            })

            if mark_seen:
                message.UnRead = False
                message.Save()

            processed_count += 1

        except Exception as e:
            print(f"[EmailMonitor] Error processing message: {e}")
            continue

    print(f"[EmailMonitor] Scanned {unread_count} unread email(s), "
          f"found {processed_count} with supported attachments.")
    return results


if __name__ == "__main__":
    # Quick test — run this file directly to verify Outlook connection
    print("Testing Outlook connection...")
    emails = fetch_new_emails(mark_seen=False)  # mark_seen=False = safe test mode
    if emails:
        print(f"\nFound {len(emails)} email(s) with attachments:")
        for e in emails:
            print(f"  Subject : {e['subject']}")
            print(f"  From    : {e['sender']}")
            print(f"  To      : {', '.join(e['to_addresses'])}")
            print(f"  Files   : {[a['filename'] for a in e['attachments']]}")
            print()
    else:
        print("No unread emails with attachments found.")
        print("(This is normal if your inbox is empty or all emails are already read.)")
