"""
Agent 1: Email Monitor (Outlook COM version)
Reads directly from Outlook Classic on this machine — no IMAP credentials needed.
Requires: pywin32, Outlook installed and open with the orders inbox loaded.

Attachments are saved to:
  data/inbox/YYYYMM/{VendorName}/{timestamp}_{filename}

All emails with supported attachments are downloaded and organized by vendor.
Vendor is determined by the To: address domains or sender domain.
"""

import os
import re
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

_SCRIPT_DIR  = Path(__file__).resolve().parent.parent
STORAGE_ROOT = os.getenv("STORAGE_ROOT", str(_SCRIPT_DIR / "data"))
SUPPORTED_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".pdf"}
TARGET_INBOX = os.getenv("OUTLOOK_INBOX", None)

# ── Vendor identification ─────────────────────────────────────────────────────
# Maps email domains to folder names.
# Checked against both To: addresses and the sender's domain.
DOMAIN_TO_VENDOR = {
    # RWC
    "rwcsupply.com":        "RWC",
    "rwc.org":              "RWC",
    # L&W Supply
    "lwsupply.com":         "LW",
    "billtrust.com":        "LW",        # L&W AND White Cap invoices via billtrust
    # Sherwin-Williams
    "sherwin.com":          "Sherwin-Williams",
    "sherwin-williams.com": "Sherwin-Williams",
    # Builders FirstSource
    "bldr.com":             "Builders-FirstSource",
    # White Cap
    "whitecap.com":         "White-Cap",
    # FBM
    "myfbm.com":            "FBM",
    # Southwest Mobile Storage
    "swmobilestorage.com":  "SW-Mobile-Storage",
    # J&J Sand — handled by exact address check above due to gmail domain
    # Internal / supervisors
    "integrityllctuc.com":  "Internal",
    "logisticconsultants.com": "Internal",
    "txexterior.com":       "Internal",
    "gmail.com":            "Internal",  # tsp.iws@gmail.com forwards
    "utdallas.edu":         "Internal",  # remove after testing
}


def _safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


def _get_inbox(outlook_ns):
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
    return outlook_ns.GetDefaultFolder(6)


def _extract_to_addresses(message) -> list[str]:
    addresses = []
    try:
        for i in range(1, message.Recipients.Count + 1):
            recipient = message.Recipients.Item(i)
            addr = recipient.Address or ""
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


# ── Inventory filename patterns ───────────────────────────────────────────────
# Files matching these patterns are inventory/report files, not PO forms.
# They get routed to the Inventory folder regardless of sender/vendor.
INVENTORY_PATTERNS = [
    r"(?i)\beom\b",          # End of Month: "EOM March 2026", "BK SBR EOM 03.2026"
    r"(?i)\binventory\b",    # any file with "inventory" in the name
    r"(?i)\bstatement\b",    # monthly statements
    r"(?i)pricing.effect",   # pricing sheets e.g. "RWC Pricing Effect 12.1.25"
    r"(?i)price.quote",      # quote documents
]


def _is_inventory_file(filename: str) -> bool:
    """Return True if the filename matches known inventory/report patterns."""
    return any(re.search(pattern, filename) for pattern in INVENTORY_PATTERNS)
    """
    Determine vendor folder name.
    Order of priority:
    1. Exact sender address (handles gmail senders like J&J Sand)
    2. To: address domains (most reliable for POs)
    3. Sender domain (catches invoices sent directly)
    4. 'Other' if nothing matches
    """
    sender_lower = sender.lower()

    # ── 1. Exact sender address matches ──────────────────────────────────────
    if "jjsandtucson@gmail.com" in sender_lower:
        return "JJ-Sand"

    # ── 2. To: address domains ────────────────────────────────────────────────
    for addr in to_addresses:
        domain = addr.split("@")[-1].lower() if "@" in addr else ""
        for vendor_domain, vendor_name in DOMAIN_TO_VENDOR.items():
            if vendor_domain in domain:
                return vendor_name

    # ── 3. Sender domain fallback ─────────────────────────────────────────────
    sender_domain = sender_lower.split("@")[-1] if "@" in sender_lower else ""
    for vendor_domain, vendor_name in DOMAIN_TO_VENDOR.items():
        if vendor_domain in sender_domain:
            return vendor_name

    return "Other"


def _vendor_folder(vendor_name: str) -> Path:
    """
    Return the folder path: data/inbox/YYYYMM/{VendorName}/
    Creates the folder if it doesn't exist.
    """
    month  = datetime.now().strftime("%Y%m")
    folder = Path(STORAGE_ROOT) / "inbox" / month / vendor_name
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def fetch_new_emails(mark_seen: bool = True) -> list[dict]:
    """
    Connect to Outlook, fetch all unread emails with supported attachments.
    Saves files organized by vendor and month.
    Returns all emails — vendor filtering for PO validation happens in Agent 2.
    """
    try:
        import win32com.client
    except ImportError:
        print("[EmailMonitor] ERROR: pywin32 not installed. Run: pip install pywin32")
        return []

    results = []

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns      = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"[EmailMonitor] Could not connect to Outlook: {e}")
        print("  Make sure Outlook is open and signed in.")
        return []

    try:
        inbox = _get_inbox(ns)
    except Exception as e:
        print(f"[EmailMonitor] Could not access inbox: {e}")
        return []

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    unread_count = processed_count = 0

    for message in messages:
        try:
            if message.UnRead is False:
                continue

            unread_count += 1

            if message.Attachments.Count == 0:
                if mark_seen:
                    message.UnRead = False
                    message.Save()
                continue

            subject      = message.Subject or ""
            sender       = message.SenderEmailAddress or message.SenderName or ""
            received_at  = str(message.ReceivedTime)
            uid          = str(message.EntryID)
            to_addresses = _extract_to_addresses(message)

            # Determine vendor for folder routing
            vendor   = _detect_vendor(to_addresses, sender)
            save_dir = _vendor_folder(vendor)

            attachments = []
            for i in range(1, message.Attachments.Count + 1):
                attachment = message.Attachments.Item(i)
                filename   = _safe_filename(attachment.FileName or f"attachment_{i}")
                ext        = Path(filename).suffix.lower()

                if ext not in SUPPORTED_EXTENSIONS:
                    print(f"[EmailMonitor] Skipping: {filename}")
                    continue

                # Route inventory/report files to their own folder
                if _is_inventory_file(filename):
                    file_folder = _vendor_folder("Inventory")
                    folder_label = "Inventory"
                else:
                    file_folder  = save_dir
                    folder_label = vendor

                ts        = datetime.now().strftime("%Y%m%d_%H%M%S")
                save_name = f"{ts}_{filename}"
                save_path = file_folder / save_name

                attachment.SaveAsFile(str(save_path))
                print(f"[EmailMonitor] [{folder_label}] {save_name}")

                attachments.append({
                    "filename":  filename,
                    "filepath":  str(save_path),
                    "extension": ext,
                })

            if not attachments:
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
                "vendor":       vendor,
                "attachments":  attachments,
            })

            if mark_seen:
                message.UnRead = False
                message.Save()

            processed_count += 1

        except Exception as e:
            print(f"[EmailMonitor] Error processing message: {e}")
            continue

    print(f"[EmailMonitor] {unread_count} unread | {processed_count} with attachments saved")
    return results


if __name__ == "__main__":
    print("Testing Outlook connection...")
    emails = fetch_new_emails(mark_seen=False)
    if emails:
        print(f"\nFound {len(emails)} email(s) with attachments:")
        for e in emails:
            print(f"  [{e['vendor']}] {e['subject']}")
            print(f"    Files: {[a['filename'] for a in e['attachments']]}")
            print()
    else:
        print("No unread emails with attachments found.")
