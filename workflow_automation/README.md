# MCP Purchase Order Validator

A lightweight AI-powered application built using the Model Context Protocol (MCP) that detects discrepancies between purchase orders (POs) and invoices.

This project demonstrates how an AI assistant (e.g., Claude Desktop) can interact with structured financial data and perform validation tasks in real time.

---

## Problem

In accounts payable workflows, invoice discrepancies are common and often require manual review. This process is:

- Time-consuming  
- Error-prone  
- Difficult to scale  

---

## Solution

This MCP server enables an AI assistant to:

- Store purchase order data  
- Record invoice submissions  
- Automatically compare values  
- Flag discrepancies instantly  

---

## Features

| Tool | Description |
|------|------------|
| `add_po` | Store a purchase order (PO number, vendor, amount) |
| `add_invoice` | Record an invoice tied to a PO |
| `check_po_status` | Compare PO vs invoice and detect discrepancies |

---

## How It Works

1. A PO is created with an expected amount  
2. An invoice is submitted  
3. The AI compares both values  
4. The system returns:
   - Match  
   - No invoice  
   - Discrepancy detected  

---

## Example

**Input:**

## Demo

Example workflow:

1. Add a purchase order:
   add_po("PO123", "Vendor A", 1000)

2. Add an invoice:
   add_invoice("PO123", 1200)

3. Check status:
   check_po_status("PO123")

Output:
Discrepancy detected: $200.00