# PO Automation System
**Logistic Consultants — Orders Inbox Processor**

Automatically monitors the shared orders inbox, parses purchase orders and invoices,
validates for duplicates and vendor mismatches, files documents, and sends alerts.

---

## Architecture

```
run_automation.py (orchestrator)
│
├── Agent 1: email_monitor.py    — IMAP fetch, attachment extraction
├── Agent 2: document_parser.py  — XLS PO parsing, PDF invoice parsing (Claude)
├── Agent 3: validator.py        — Duplicate check, vendor mismatch, logging
└── Agent 4: filer.py            — Local file storage, alert emails
```

### Folder structure created automatically
```
data/
  inbox/           ← raw attachments as received
  processed/
    YYYYMM/
      POs/         ← clean PO files
      Invoices/    ← clean invoices
  flagged/         ← anything with a duplicate/mismatch flag
  archive/
  po_tracking.csv  ← master log (replaces the Excel PO List sheet)
```

---

## Setup

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure credentials
```bash
cp .env.example .env
# Edit .env with your IMAP credentials and API keys
```

#### Email options:
- **Option A (preferred):** Direct IMAP to Actionweb hosting
  - Log into cPanel → Email Accounts → confirm IMAP is enabled
  - Set `IMAP_HOST=mail.logisticconsultants.com`
- **Option B (fallback):** Gmail IMAP
  - Enable 2FA on Google account
  - Generate an App Password: myaccount.google.com → Security → App Passwords
  - Set `IMAP_HOST=imap.gmail.com` and use the App Password

### 3. Add vendor domain mappings
Edit `agents/document_parser.py` → `VENDOR_DOMAIN_MAP` to add/correct vendor domains:
```python
VENDOR_DOMAIN_MAP = {
    "rwc":   "rwcsupply.com",
    "l&w":   "lwsupply.com",
    ...
}
```

---

## Running

### Manual run (one-time check):
```bash
python run_automation.py
```

### Dry run (no files moved, no emails sent):
```bash
python run_automation.py --dry-run
```

### Dashboard UI:
```bash
streamlit run app.py
```

---

## Scheduling (Windows Task Scheduler)

1. Open Task Scheduler → Create Basic Task
2. Name: "PO Inbox Monitor"
3. Trigger: Daily, repeat every 15 minutes
4. Action: Start a program
   - Program: `python`
   - Arguments: `run_automation.py`
   - Start in: `C:\path\to\po-automation`
5. Check "Run whether user is logged on or not"

---

## PO Form Cell Reference (from VBA macros)

| Field         | Lath PO  | Stucco PO |
|---------------|----------|-----------|
| PO Number     | I12      | F12       |
| Order Date    | I6       | H6        |
| Delivery Date | I7       | H7        |
| Address       | H10      | G11       |
| Track/Vendor  | B7       | B7        |
| Release       | B8       | B8        |
| Lot           | D9       | D9        |
| Type detect   | A12 (contains "lath" or "stucco") |
| Items         | rows 14–46, col A=desc col B=qty | rows 14–29 |
| Supervisor    | Last 2 chars of PO# (BK or AN)  |

---

## Future: Merging with Purchase_Order_Compiler.xlsm

The `po_tracking.csv` log mirrors the "Purchase Order List" sheet columns exactly.
A future step can either:
- Export CSV → import to the compiler's "Purchase Order List" sheet
- Or replace the compiler entirely by expanding `run_automation.py` to write XLSX directly
  using `openpyxl`, which would make the VBA macros unnecessary.

---

## Vendor Mismatch Detection

The system flags when the email's `To:` address domain doesn't match
the vendor named in the PO form.

**Example:** Form says "RWC" but email was sent to `vendor@LWSupply.com`
→ logged as `VENDOR_MISMATCH`, filed to `data/flagged/`, alert email sent.
