# Accounts-Payable-Agent

**BUAN 6V99.S01 S26 — Agentic AI | Professor Antonio Paes**
**Danielle Beagle | University of Texas at Dallas**

---

## Overview

An automated multi-agent system that monitors a shared company email inbox, processes purchase order and invoice attachments, validates for errors, and organizes files — replacing a manual review workflow that was prone to duplicate PO numbers, vendor mismatches, and missing documentation.


## Problem Statement

Supervisors submit purchase orders as Excel attachments to a shared orders inbox (`orders@logisticconsultants.com`), then copy vendors directly. This process creates recurring errors:

- **Duplicate PO numbers** — supervisors reuse old forms without updating the PO number
- **Vendor mismatches** — the form names one vendor (e.g. RWC) but the email is sent to a different vendor (e.g. L&W Supply)
- **Missing invoices** — invoices arrive without a corresponding PO on file
- **Manual overhead** — all of the above required manual detection and correction

---

## Solution Architecture

The system uses four specialized agents orchestrated by a central runner, plus a Streamlit dashboard for manual review.

```
run_automation.py  (orchestrator — scheduled via Task Scheduler)
│
├── Agent 1: email_monitor.py
│   └── Connects via IMAP to shared inbox
│   └── Fetches unread emails with attachments
│   └── Saves attachments locally, extracts To:/From: metadata
│
├── Agent 2: document_parser.py
│   └── .xls/.xlsx → extracts PO number, vendor, dates, line items
│   └── .pdf → uses Claude AI to extract structured invoice data
│   └── Detects PO type (Lath / Stucco) from form cell A12
│
├── Agent 3: validator.py
│   └── Checks PO number against tracking log for duplicates
│   └── Compares vendor on form vs. vendor implied by email To: domain
│   └── Matches invoices to existing POs
│   └── Logs all results to po_tracking.csv
│
└── Agent 4: filer.py
    └── Routes clean files → data/processed/YYYYMM/
    └── Routes flagged files → data/flagged/
    └── Sends alert emails for duplicates, mismatches, and invoice issues
```

---

## Key Features

| Feature | Description |
|---|---|
| Duplicate detection | Flags any PO number already present in the tracking log |
| Vendor mismatch detection | Compares vendor named on the Excel form vs. the email recipient domain |
| Invoice matching | Links PDF invoices to existing POs by reference number |
| Alert emails | Automatic notification sent when errors are detected |
| Organized file storage | Files sorted into monthly folders by type and status |
| Streamlit dashboard | Manual review UI with filtering, detail view, and CSV export |
| Dry-run mode | Test parsing and validation without moving files or sending emails |

---

## PO Form Structure

The parser mirrors the exact cell references used in the existing `Purchase_Order_Compiler.xlsm` VBA macros for compatibility.

| Field | Lath PO | Stucco PO |
|---|---|---|
| PO Number | I12 | F12 |
| Order Date | I6 | H6 |
| Delivery Date | I7 | H7 |
| Address | H10 | G11 |
| Vendor / Track | B7 | B7 |
| Release | B8 | B8 |
| Lot | D9 | D9 |
| Type detection | A12 contains "lath" or "stucco" |
| Supervisor code | Last 2 chars of PO# — `BK` or `AN` |
| Line items | Rows 14–46 (Lath) / 14–29 (Stucco) |

---

## Tech Stack

| Layer | Technology |
|---|---|
| Language | Python 3.11+ |
| Email access | `imaplib` — direct IMAP to Actionweb hosting or Gmail fallback |
| XLS parsing | `xlrd` — reads legacy `.xls` PO forms |
| PDF parsing | `pdfplumber` + Anthropic Claude API |
| AI extraction | Claude claude-sonnet-4-20250514 via Anthropic Python SDK |
| Dashboard | Streamlit |
| Data logging | CSV (`po_tracking.csv`) — compatible with Excel PO List sheet |
| Scheduling | Windows Task Scheduler |

---

## Project Structure

```
Accounts-Payable-Agent/
├── po-automation/
│   ├── run_automation.py       # Orchestrator — run this on a schedule
│   ├── app.py                  # Streamlit dashboard
│   ├── .env.example            # Credential config template
│   ├── requirements.txt
│   ├── README.md               # Setup and scheduling instructions
│   └── agents/
│       ├── email_monitor.py    # Agent 1
│       ├── document_parser.py  # Agent 2
│       ├── validator.py        # Agent 3
│       └── filer.py            # Agent 4
└── workflow_automation/        # Original n8n workflow prototype (reference)
```

---

## Setup

```bash
cd po-automation
pip install -r requirements.txt
cp .env.example .env
# Fill in IMAP credentials and Anthropic API key in .env
```

**Email options:**
- **Option A (preferred):** Direct IMAP to `mail.logisticconsultants.com` — no forwarding lag
- **Option B (fallback):** Gmail IMAP via `imap.gmail.com` using a Google App Password

**Run manually:**
```bash
python run_automation.py --dry-run   # test without moving files
python run_automation.py             # live run
streamlit run app.py                 # launch dashboard
```

**Schedule on Windows (every 15 min):**
Task Scheduler → New Task → Trigger: repeat every 15 min → Action: `python run_automation.py`

---

## Data Flow

```
Shared inbox (orders@logisticconsultants.com)
    ↓ IMAP fetch (Agent 1)
Raw attachments saved to data/inbox/
    ↓ Parse (Agent 2)
Structured data: PO number, vendor, dates, items
    ↓ Validate (Agent 3)
Status: CLEAN | DUPLICATE | VENDOR_MISMATCH | INVOICE_ISSUE
    ↓ File + Notify (Agent 4)
data/processed/YYYYMM/   ← clean documents
data/flagged/            ← anything that needs review
Alert email              ← sent if errors detected
po_tracking.csv          ← master log of all POs
```

---

## Relation to Existing Tools

The `po_tracking.csv` output mirrors the column structure of the **Purchase Order List** sheet in `Purchase_Order_Compiler.xlsm`. This was intentional — the automation can feed the existing Excel workflow, or eventually replace it entirely as confidence in the system grows.
