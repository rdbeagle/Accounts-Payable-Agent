# Accounts-Payable-Agent
Project for Agentic AI Course BUAN 6v99.s01 S26 with professor Antonio Paes

# Purpose
This is an automation agent for the shared purchase order inbox that helps reduce errors in the ordering process. 

Supervisors submit purchase orders as Excel attachments to the orders email, but they sometimes reuse old forms, use duplicate PO numbers, or forget to update vendor information, which creates confusion and accounting issues.

The system we are building identifies the vendor based on who the email was sent to, extracts the purchase order number from the file, checks for mismatches or duplicate numbers, and organizes attachments automatically by vendor and date into our local files to be later compiled via the excel macros. 

It also processes incoming vendor invoices, separates them if they arrive in batches, matches them to existing purchase orders, and sends alerts if an invoice arrives without a corresponding order or if there’s a numbering conflict.

Additionally when errors such as receiving an invoice that does not have a corresponding PO on file, and automatic email will be sent to the supervisor detailing what error was found and what corrective actions need to be taken.
