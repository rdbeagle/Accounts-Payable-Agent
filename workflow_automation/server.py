from mcp.server.fastmcp import FastMCP

mcp = FastMCP("po-validator")

purchase_orders = {}
invoices = {}

@mcp.tool()
def add_po(po_number: str, vendor: str, amount: float):
    purchase_orders[po_number] = {
        "vendor": vendor,
        "amount": amount
    }
    return f"PO {po_number} added."

@mcp.tool()
def add_invoice(po_number: str, amount: float):
    invoices[po_number] = amount
    return f"Invoice for {po_number} added."

@mcp.tool()
def check_po_status(po_number: str):
    if po_number not in purchase_orders:
        return "PO not found"
    
    if po_number not in invoices:
        return "No invoice submitted"
    
    po_amount = purchase_orders[po_number]["amount"]
    invoice_amount = invoices[po_number]
    
    if po_amount == invoice_amount:
        return "Match"
    else:
        diff = invoice_amount - po_amount
        return f"Discrepancy detected: ${diff}"

if __name__ == "__main__":
    mcp.run(transport="stdio")