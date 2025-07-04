#!/usr/bin/env python3
"""
QuickBooks Online – Bulk Invoice Import Script
=============================================

Version 2025‑06‑25 e‑qty‑rate‑optional‑3

WHAT'S NEW (patch‑3)
--------------------
* **True omission of Qty / UnitPrice / TaxInclusiveAmt when blank.** `python‑quickbooks`
  sets these to **0** by default (see snippet on GitHub showing `self.UnitPrice = 0`
  and `self.Qty = 0`). We now *force them back to `None`* right after object
  instantiation and only overwrite when an explicit value is present.  This keeps
  your payload fully compliant with Intuit's "optional" rules.
* Added helper `_apply_qty_rate(detail, qty, rate)` to tidy the logic.
* Bumped version tag & changelog accordingly. All CLI flags unchanged.
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import sys
from datetime import datetime
from typing import Any, Dict, List, Optional

from dotenv import load_dotenv
from intuitlib.client import AuthClient
from quickbooks import QuickBooks
from quickbooks.objects import (
    Customer,
    Invoice,
    Item,
    SalesItemLine,
    SalesItemLineDetail,
    # SalesTerm does not exist in python-quickbooks; use Term
)
try:
    from quickbooks.objects import Term
except ImportError:
    Term = None

# Load environment variables from .env file
load_dotenv()

# ---------------------------------------------------------------------------
# Configuration – Load from environment variables
# ---------------------------------------------------------------------------
CONFIG: Dict[str, Optional[str] | bool] = {
    "CLIENT_ID": os.getenv("CLIENT_ID"),  # Intuit Developer – Client ID
    "CLIENT_SECRET": os.getenv("CLIENT_SECRET"),  # Intuit Developer – Client Secret
    "SANDBOX": os.getenv("SANDBOX", "true").lower() == "true",  # Flip to False for production
    "REDIRECT_URI": os.getenv("REDIRECT_URI", "https://developer.intuit.com/v2/OAuth2Playground/RedirectUrl"),
    "ACCESS_TOKEN": None,
    "REFRESH_TOKEN": None,
    "REALM_ID": None,
}

def validate_environment() -> bool:
    """Validate that required environment variables are set."""
    required_vars = ["CLIENT_ID", "CLIENT_SECRET"]
    missing_vars = []
    
    for var in required_vars:
        if not os.getenv(var):
            missing_vars.append(var)
    
    if missing_vars:
        print("❌ Missing required environment variables:")
        for var in missing_vars:
            print(f"   - {var}")
        print("\nPlease create a .env file with your credentials:")
        print("1. Copy .env.example to .env")
        print("2. Fill in your actual CLIENT_ID and CLIENT_SECRET values")
        return False
    
    return True

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def parse_date(date_str: str, fmt: str = "%m/%d/%Y") -> str:
    """Convert *MM/DD/YYYY* strings to ISO *YYYY‑MM‑DD* for QBO."""
    return datetime.strptime(date_str.strip(), fmt).date().isoformat()


def to_float_or_none(val: str) -> Optional[float]:
    val = val.strip()
    if not val:
        return None
    try:
        return float(val)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# CSV ingestion
# ---------------------------------------------------------------------------

def read_invoices(csv_path: str) -> Dict[str, Any]:
    """Parse a CSV (exported from Excel/Sheets) into a nested dict keyed by InvoiceNo."""
    invoices: Dict[str, Any] = {}
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, skipinitialspace=True)
        reader.fieldnames = [h.strip() for h in reader.fieldnames]
        for row in reader:
            inv_no = row["InvoiceNo"].strip()
            if inv_no not in invoices:
                invoices[inv_no] = {
                    "Customer": row["Customer"].strip(),
                    "InvoiceDate": parse_date(row["InvoiceDate"]),
                    "DueDate": parse_date(row["DueDate"]),
                    "CustomerMemo": row.get("CustomerMemo", "").strip(),
                    "Terms": row.get("Terms", "").strip(),  # Add Terms
                    "LineItems": [],
                }
            invoices[inv_no]["LineItems"].append(
                {
                    "Item": row.get("Item(Product/Service)", "").strip(),
                    "Description": row.get("ItemDescription", "").strip(),
                    "Quantity": to_float_or_none(row.get("ItemQuantity", "")),
                    "Rate": to_float_or_none(row.get("ItemRate", "")),
                    "Amount": float(row.get("ItemAmount", 0) or 0),
                }
            )
    return invoices

# ---------------------------------------------------------------------------
# OAuth helpers (unchanged)
# ---------------------------------------------------------------------------

def load_tokens() -> bool:
    if os.path.exists("qb_tokens.json"):
        with open("qb_tokens.json", "r", encoding="utf-8") as f:
            tokens = json.load(f)
            CONFIG.update(
                {
                    "ACCESS_TOKEN": tokens.get("access_token"),
                    "REFRESH_TOKEN": tokens.get("refresh_token"),
                    "REALM_ID": tokens.get("realm_id"),
                }
            )
            return True
    return False


def setup_oauth() -> bool:
    print("=== OAuth Setup Required ===")
    print(
        "Run the OAuth playground (https://appcenter.intuit.com/playground) "
        "and save the resulting tokens to qb_tokens.json before continuing."
    )
    return False

# ---------------------------------------------------------------------------
# QuickBooks helpers (unchanged)
# ---------------------------------------------------------------------------

def initialize_quickbooks_client() -> Optional[QuickBooks]:
    if not CONFIG["ACCESS_TOKEN"] or not CONFIG["REALM_ID"]:
        print("❌ Missing OAuth tokens. Please run the OAuth setup.")
        return None
    try:
        auth_client = AuthClient(
            client_id=CONFIG["CLIENT_ID"],
            client_secret=CONFIG["CLIENT_SECRET"],
            redirect_uri=CONFIG["REDIRECT_URI"],
            environment="sandbox" if CONFIG["SANDBOX"] else "production",
            access_token=CONFIG["ACCESS_TOKEN"],
            refresh_token=CONFIG["REFRESH_TOKEN"],
        )
        return QuickBooks(
            auth_client=auth_client,
            refresh_token=CONFIG["REFRESH_TOKEN"],
            company_id=CONFIG["REALM_ID"],
            sandbox=CONFIG["SANDBOX"],
        )
    except Exception as exc:
        print(f"❌ Failed to initialize QuickBooks client: {exc}")
        return None

# ---------------------------------------------------------------------------
# Customer & Item helpers (unchanged)
# ---------------------------------------------------------------------------

def find_or_create_customer(client: QuickBooks, name: str) -> Optional[Customer]:
    try:
        customers = Customer.filter(DisplayName=name, qb=client)
        if customers:
            return customers[0]
        print(f"❌ Customer '{name}' not found in QuickBooks. Please create the customer first.")
        return None
    except Exception as exc:
        print(f"❌ Error with customer '{name}': {exc}")
        return None


def find_or_create_item(client: QuickBooks, item_name: str) -> Optional[Item]:
    try:
        items = Item.filter(Name=item_name, qb=client)
        if items:
            return items[0]
        print(f"❌ Item '{item_name}' not found in QuickBooks. Please create the item first.")
        return None
    except Exception as exc:
        print(f"❌ Error with item '{item_name}': {exc}")
        return None

def find_sales_term_by_name(client: QuickBooks, term_name: str):
    """Find a Term by name. Returns the object or None."""
    if not Term:
        print("❌ Term object not available in python-quickbooks.")
        return None
    try:
        terms = Term.filter(Name=term_name, qb=client)
        if terms:
            return terms[0]
        print(f"❌ Term '{term_name}' not found in QuickBooks. Please create the term first.")
        return None
    except Exception as exc:
        print(f"❌ Error with term '{term_name}': {exc}")
        return None

# ---------------------------------------------------------------------------
# Utility to strip default zeros
# ---------------------------------------------------------------------------

def _apply_qty_rate(detail: SalesItemLineDetail, qty: Optional[float], rate: Optional[float]):
    """Overwrite SDK's default 0 values so blank CSV cols stay truly blank."""
    # Reset defaults introduced by python‑quickbooks
    detail.Qty = None  # type: ignore
    detail.UnitPrice = None  # type: ignore
    detail.TaxInclusiveAmt = None  # type: ignore

    # Now apply real values if provided
    if qty is not None:
        detail.Qty = qty  # type: ignore
    if rate is not None:
        detail.UnitPrice = rate  # type: ignore

# ---------------------------------------------------------------------------
# Core Invoice creation routine
# ---------------------------------------------------------------------------

def invoice_number_exists(client: QuickBooks, doc_number: str) -> bool:
    """Check if an invoice with the given DocNumber exists in QBO."""
    try:
        # The filter method may not be documented, but works for other objects
        results = Invoice.filter(DocNumber=doc_number, qb=client)
        return bool(results)
    except Exception as exc:
        print(f"❌ Error checking for existing invoice number '{doc_number}': {exc}")
        return False

def create_quickbooks_invoice(
    client: QuickBooks,
    data: Dict[str, Any],
    inv_no: str,
    *,
    debug_json: bool = False,
    only_required: bool = False,
    auto_fill_qty_rate: bool = False,
) -> bool:
    try:
        # Check if invoice number already exists
        if invoice_number_exists(client, inv_no):
            print(f"❌ Invoice number '{inv_no}' already exists in QBO. Skipping import.")
            return False
        customer = find_or_create_customer(client, data["Customer"])
        if not customer:
            return False
        invoice = Invoice()
        invoice.CustomerRef = customer.to_ref()  # type: ignore
        invoice.TxnDate = data["InvoiceDate"]
        invoice.DocNumber = inv_no  # type: ignore
        if not only_required:
            invoice.DueDate = data["DueDate"]
            if data.get("CustomerMemo"):
                invoice.CustomerMemo = {"value": data["CustomerMemo"]}  # type: ignore
            # Set SalesTermRef if Terms is present
            if data.get("Terms"):
                term = find_sales_term_by_name(client, data["Terms"])
                if term:
                    invoice.SalesTermRef = term.to_ref()  # type: ignore

        lines: List[SalesItemLine] = []
        for row in data["LineItems"]:
            item_name = row.get("Item")
            if not item_name:
                continue
            item = find_or_create_item(client, item_name)
            if not item:
                continue

            qty = row.get("Quantity")
            rate = row.get("Rate")
            amount = row["Amount"]

            # Legacy mode: auto‑fill Qty/Rate when blank
            if auto_fill_qty_rate:
                if qty is None and rate is None:
                    qty = 1
                    rate = amount
                elif qty is None:
                    qty = 1 if rate == 0 else round(amount / rate, 4)
                elif rate is None:
                    rate = 0 if qty == 0 else round(amount / qty, 4)

            detail = SalesItemLineDetail()
            detail.ItemRef = item.to_ref()  # type: ignore
            _apply_qty_rate(detail, qty, rate)

            line = SalesItemLine()
            if not only_required and row.get("Description"):
                line.Description = row["Description"]
            line.Amount = amount
            line.SalesItemLineDetail = detail  # type: ignore
            lines.append(line)

        if not lines:
            print("⚠️  Invoice skipped – no valid line items detected.")
            return False

        invoice.Line = lines

        if debug_json:
            raw_json = invoice.to_json()
            pretty_json = json.dumps(json.loads(raw_json), indent=2)
            print("📤 JSON payload just before save:")
            print(pretty_json)

        invoice.save(qb=client)
        print(f"✅ Created invoice {inv_no} (QBO Id: {invoice.Id})")
        return True
    except Exception as exc:
        print(f"❌ Error creating invoice: {exc}")
        return False

# ---------------------------------------------------------------------------
# Batch processing & CLI (unchanged)
# ---------------------------------------------------------------------------

def process_invoices(
    client: QuickBooks,
    invoices_data: Dict[str, Any],
    *,
    debug_json: bool = False,
    only_required: bool = False,
    auto_fill_qty_rate: bool = False,
) -> None:
    total = len(invoices_data)
    success = 0
    print(f"\n📋 Processing {total} invoices…")
    for inv_no, data in invoices_data.items():
        print(f"\n🔄 Processing Invoice {inv_no}…")
        if create_quickbooks_invoice(
            client,
            data,
            inv_no,
            debug_json=debug_json,
            only_required=only_required,
            auto_fill_qty_rate=auto_fill_qty_rate,
        ):
            success += 1
        else:
            print(f"❌ Failed to create Invoice {inv_no}")
    print(f"\n📊 Summary: {success}/{total} invoices processed successfully")


# ---------------------------------------------------------------------------
# Main entry‑point (unchanged)
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Bulk‑import invoices into QBO")
    parser.add_argument("--debug-json", action="store_true", help="Print JSON payload before save")
    parser.add_argument("--only-required", action="store_true", help="Send only mandatory fields")
    parser.add_argument(
        "--auto-fill-qty-rate",
        action="store_true",
        help="Back‑fill Qty/Rate when blank (legacy behaviour)",
    )
    args = parser.parse_args()

    print("=== QuickBooks Invoice Import Script ===")

    # Validate environment variables
    if not validate_environment():
        sys.exit(1)

    if not load_tokens() and not setup_oauth():
        sys.exit(1)

    csv_file = "invoices.csv"
    if not os.path.exists(csv_file):
        print(f"❌ CSV file '{csv_file}' not found")
        sys.exit(1)

    print(f"📄 Loading invoices from {csv_file}…")
    invoices = read_invoices(csv_file)
    print(f"✅ Loaded {len(invoices)} invoices from CSV")

    client = initialize_quickbooks_client()
    if not client:
        sys.exit(1)

    process_invoices(
        client,
        invoices,
        debug_json=args.debug_json,
        only_required=args.only_required,
        auto_fill_qty_rate=args.auto_fill_qty_rate,
    )

    print("\n🏁 Script completed!")


if __name__ == "__main__":
    main()
