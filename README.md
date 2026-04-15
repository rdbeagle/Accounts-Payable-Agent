# IWS Accounts Payable Automation Agent

An agentic AI system that automates purchase order processing for **Integrity Wall Systems** / **Logistic Consultants**, a construction subcontracting company based in Tucson, AZ.

Built for BUAN 6V99.S01 — Agentic AI, University of Texas at Dallas.

---

## The Problem

Integrity Wall Systems sends 30–60 purchase orders per week via email to vendors (RWC Supply, L&W Supply, and others). The orders inbox (`orders@logisticconsultants.com`) receives POs, invoices, inventory reports, and other documents mixed together. Previously, an office manager manually reviewed every email, filed documents, and checked for errors — a slow, error-prone process.

**This system automates the entire pipeline.**

---

## What It Does

- **Monitors the Outlook inbox** in real time via Outlook COM (no IMAP credentials needed)
- **Parses XLS purchase order forms** — reads fixed cell addresses for both Lath and Stucco form types
- **Detects duplicates** — flags if a PO number has already been logged
- **Detects vendor mismatches** — flags if the vendor on the form doesn't match who the email was sent to
- **Matches invoices to POs** — uses Claude to extract vendor and PO reference from PDF invoices, cross-checks against the tracking log
- **Files documents** automatically into organized folders by vendor and month
- **Sends contextual alert emails** via Outlook — routed to the supervising field manager (Blake or Adam), with the office manager CC'd, written by Claude to explain the specific issue and what action to take
- **Maintains a CSV tracking log** of all processed POs
- **Displays a live dashboard** with an AI-generated inbox brief, KPI cards, flagged item alerts, full PO log with search/filter, and PO detail view with line items

---

## Architecture — Four Agents

```
Agent 1: Email Monitor (email_monitor.py)
  └── Connects to Outlook Classic via win32com
  └── Downloads attachments, routes to vendor folders by sender domain
  └── Skips images, non-supported files, and inventory reports

Agent 2: Document Parser (document_parser.py)
  └── XLS: reads fixed cell addresses for Lath and Stucco PO forms
  └── PDF: uses Claude to extract invoice fields (vendor, PO reference, invoice number)

Agent 3: Validator (validator.py)
  └── Checks PO number against tracking log (duplicate detection)
  └── Checks vendor on form vs. email recipient (mismatch detection)
  └── Cross-checks invoices against tracking log for PO existence and vendor match

Agent 4: Filer & Notifier (filer.py)
  └── Routes clean POs to processed/YYYYMM/POs/
  └── Routes flagged documents to flagged/
  └── Uses Claude to write contextual alert emails
  └── Sends alerts via Outlook COM from orders@logisticconsultants.com
  └── Routes to Blake (BK) or Adam (AN) based on supervisor code in PO number
```

---

## AI / LLM Features

Claude is used in three places:

1. **Invoice PDF parsing** — extracts vendor name, PO reference number, and invoice number from unstructured PDF text
2. **Alert email authoring** — writes a plain English explanation of what went wrong and what action to take, specific to each flag
3. **AI Inbox Brief** — generates a 3-4 sentence dashboard summary of the current inbox state, flagged items, and most important action needed

---

## Tech Stack

- **Python 3.13**
- **pywin32** — Outlook COM automation (email reading + sending)
- **xlrd** — XLS purchase order parsing
- **pdfplumber** — PDF text extraction
- **anthropic** — Claude API (claude-haiku-4-5)
- **streamlit** — Dashboard UI
- **pandas** — CSV tracking log
- **python-dotenv** — Environment configuration

---

## Project Structure

```
Accounts_Payable/
  agents/
    __init__.py
    email_monitor.py      # Agent 1 — Outlook inbox monitor
    document_parser.py    # Agent 2 — XLS + PDF parser
    validator.py          # Agent 3 — Duplicate, mismatch, invoice checker
    filer.py              # Agent 4 — File router + Claude alert emails
  demo/
    attachments/          # Place anonymized sample files here
    emails.json           # Simulated email metadata
    README.md
  po-automation/
    data/
      inbox/              # Downloaded attachments by vendor/month
      processed/          # Clean POs filed here
      flagged/            # Flagged documents filed here
      po_tracking.csv     # Master tracking log
  app.py                  # Streamlit dashboard
  run_automation.py       # Live orchestrator (requires Outlook)
  run_demo.py             # Demo mode orchestrator (no Outlook needed)
  requirements.txt
  .env.example
```

---

## Setup

### 1. Clone the repo
```bash
git clone https://github.com/rdbeagle/Accounts-Payable-Agent.git
cd Accounts-Payable-Agent
```

### 2. Create virtual environment
```bash
python -m venv .venv
.venv\Scripts\Activate.ps1  # Windows
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Configure environment
Copy `.env.example` to `.env` and fill in your values:
```
OUTLOOK_INBOX=orders@logisticconsultants.com
STORAGE_ROOT=./po-automation/data
ANTHROPIC_API_KEY=your_key_here
EMAIL_BLAKE=blakek@integrityllctuc.com
EMAIL_ADAM=adamn@integrityllctuc.com
EMAIL_DONNA=donnam@logisticconsultants.com
ALERT_EMAIL=orders@logisticconsultants.com
```

---

## Running the System

### Live mode (requires Outlook open with inbox loaded)
```bash
python run_automation.py        # process new emails
python run_automation.py --dry-run  # test run, no files moved
```

### Demo mode (no Outlook needed — for graders and presentations)
```bash
python run_demo.py --setup      # create demo folder structure
# add anonymized XLS files to demo/attachments/
python run_demo.py              # run full pipeline on sample files
```

### Dashboard
```bash
streamlit run app.py
```

---

## Demo Mode (For Graders)

The demo mode runs the full pipeline against local sample files without requiring access to the company's Outlook inbox.

**Setup:**
1. Run `python run_demo.py --setup` to create the folder structure
2. Add anonymized PO files to `demo/attachments/`:
   - At least one Lath PO (`.xls`)
   - At least one Stucco PO (`.xls`)
   - One duplicate (same file, two entries in `emails.json`)
3. Run `python run_demo.py`
4. Run `streamlit run app.py` to view results

The demo will show: PO parsing, duplicate detection, vendor mismatch detection, file organization, and the AI dashboard brief.

---

## Business Context

**Integrity Wall Systems** is a licensed plastering and stucco subcontractor operating in the Tucson, AZ market, working on residential developments including Saddlebrooke Ranch, Quail Creek, and Arbor at Madera. Purchase orders are issued by field supervisors (Blake and Adam) directly to vendors via email, with the orders inbox CC'd for recordkeeping.

This system processes the real inbox in production and caught a live duplicate PO on its first run.

---

## Future Roadmap

- Windows Task Scheduler for automatic 15-minute runs
- J&J Sand order parser (different form format)
- Read-only dashboard accessible from office computers
- Weekly summary email (Monday morning)
- Invoice line-item matching once PO costs are entered
