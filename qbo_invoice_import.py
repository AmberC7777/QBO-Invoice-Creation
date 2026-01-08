#!/usr/bin/env python3
"""
QuickBooks Online – Bulk Invoice Import Script
=============================================
Version 2025-07-29-servicedate-fix-1
WHAT'S NEW (servicedate-fix-1)
--------------------
* **Fixed Default ServiceDate.** The script now prevents the `python-quickbooks`
  library from inserting a default unwanted `ServiceDate` (e.g., 12/31/9999) on
  line items by explicitly setting it to None.
WHAT'S NEW (desc-only-lines-2)
--------------------
* **Corrected Description-Only Lines.** Fixed a bug where description-only
  lines were being created as standard "Sales" items. The script now correctly
  sets the `DetailType` to `DescriptionOnly` for these lines.
WHAT'S NEW (token-refresh-2)
--------------------
* **Fixed Infinite Loop.** Corrected the token refresh logic to prevent an
  infinite loop. The script now retries a failed operation only once after
  a token refresh.
WHAT'S NEW (token-refresh-1)
--------------------
* **Automatic Token Refresh.** The script now detects a 401 Authentication Error,
  automatically refreshes the access token, and retries the failed operation.
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
# Import AuthorizationException for handling 401 errors
from quickbooks.exceptions import AuthorizationException
from qb_auth import refresh_access_token, TOKEN_FILE
from quickbooks.objects import (
    Customer,
    Invoice,
    Item,
    SalesItemLine,
    SalesItemLineDetail,
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
    "CLIENT_ID": os.getenv("CLIENT_ID"),
    "CLIENT_SECRET": os.getenv("CLIENT_SECRET"),
    "SANDBOX": os.getenv("SANDBOX", "true").lower() == "true",
    "REDIRECT_URI": os.getenv("REDIRECT_URI", "https://developer.intuit.com/v2/OAuth2Playground/RedirectUrl"),
    "ACCESS_TOKEN": None,
    "REFRESH_TOKEN": None,
    "REALM_ID": None,
}
def validate_environment() -> bool:
    """Validate that required environment variables are set."""
    required_vars = ["CLIENT_ID", "CLIENT_SECRET"]
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        print("❌ Missing required environment variables:")
        for var in missing_vars:
            print(f"   - {var}")
        print("\nPlease create a .env file with your credentials.")
        return False
    
    return True
# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------
def parse_date(date_str: str, fmt: str = "%m/%d/%Y") -> str:
    """Convert *MM/DD/YYYY* strings to ISO *YYYY-MM-DD* for QBO."""
    return datetime.strptime(date_str.strip(), fmt).date().isoformat()
def to_float_or_none(val: str) -> Optional[float]:
    val = val.strip()
    return float(val) if val else None
# ---------------------------------------------------------------------------
# CSV ingestion
# ---------------------------------------------------------------------------
def read_invoices(csv_path: str, encoding: str = "utf-8-sig") -> Dict[str, Any]:
    """Parse a CSV into a nested dict keyed by InvoiceNo.

    Tries to open the CSV using ``encoding`` and falls back to ``cp1252`` if
    decoding fails. If both attempts fail, a ``UnicodeDecodeError`` is raised
    suggesting that the CSV be saved with UTF-8 encoding.
    """
    invoices: Dict[str, Any] = {}

    def _parse_csv(f) -> None:
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
                    "Terms": row.get("Terms", "").strip(),
                    "LineItems": [],
                }
            invoices[inv_no]["LineItems"].append({
                "Item": row.get("Item(Product/Service)", "").strip(),
                "Description": row.get("ItemDescription", "").strip(),
                "Quantity": to_float_or_none(row.get("ItemQuantity", "")),
                "Rate": to_float_or_none(row.get("ItemRate", "")),
                "Amount": float(row.get("ItemAmount", 0) or 0),
            })

    try:
        with open(csv_path, newline="", encoding=encoding) as f:
            _parse_csv(f)
    except UnicodeDecodeError:
        fallback = "cp1252"
        try:
            with open(csv_path, newline="", encoding=fallback) as f:
                print(
                    f"⚠️ Unable to decode file with '{encoding}'. "
                    f"Retrying with '{fallback}'."
                )
                print("   If this issue persists, please save the CSV in UTF-8.")
                _parse_csv(f)
        except UnicodeDecodeError as exc:
            raise UnicodeDecodeError(
                exc.encoding,
                exc.object,
                exc.start,
                exc.end,
                f"{exc.reason}. Please save the CSV in UTF-8."
            ) from exc
    return invoices
# ---------------------------------------------------------------------------
# OAuth helpers
# ---------------------------------------------------------------------------
def load_tokens() -> bool:
    """Load tokens from ``qb_tokens.json`` into the global ``CONFIG``.

    ``TOKEN_FILE`` lives alongside the Python scripts, so loading works even if
    the script is executed from a different current working directory.
    """
    if TOKEN_FILE.exists():
        with open(TOKEN_FILE, "r", encoding="utf-8") as f:
            tokens = json.load(f)
            CONFIG.update({
                "ACCESS_TOKEN": tokens.get("access_token"),
                "REFRESH_TOKEN": tokens.get("refresh_token"),
                "REALM_ID": tokens.get("realm_id"),
            })
            return True
    return False


def setup_oauth() -> bool:
    """Guides user to perform initial OAuth setup."""
    print("=== OAuth Setup Required ===")
    print("Run the OAuth playground (https://appcenter.intuit.com/playground) "
          "and save the resulting tokens to qb_tokens.json before continuing.")
    return False
# ---------------------------------------------------------------------------
# QuickBooks helpers
# ---------------------------------------------------------------------------
def initialize_quickbooks_client() -> Optional[QuickBooks]:
    """Initializes and returns a QuickBooks client instance."""
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
# Customer & Item helpers
# ---------------------------------------------------------------------------
def find_or_create_customer(client: QuickBooks, name: str) -> Optional[Customer]:
    """Finds a customer by name. Does not create new ones."""
    try:
        customers = Customer.filter(DisplayName=name, qb=client)
        if customers:
            return customers[0]
        print(f"❌ Customer '{name}' not found in QuickBooks. Please create it first.")
        return None
    except AuthorizationException:
        raise
    except Exception as exc:
        print(f"❌ Error with customer '{name}': {exc}")
        return None
def find_or_create_item(client: QuickBooks, item_name: str) -> Optional[Item]:
    """Finds an item by name. Does not create new ones."""
    try:
        items = Item.filter(Name=item_name, qb=client)
        if items:
            return items[0]
        print(f"❌ Item '{item_name}' not found in QuickBooks. Please create it first.")
        return None
    except AuthorizationException:
        raise
    except Exception as exc:
        print(f"❌ Error with item '{item_name}': {exc}")
        return None
def find_sales_term_by_name(client: QuickBooks, term_name: str) -> Optional[Term]:
    """Finds a Term by name."""
    if not Term:
        return None
    try:
        terms = Term.filter(Name=term_name, qb=client)
        if terms:
            return terms[0]
        print(f"❌ Term '{term_name}' not found. Please create it first.")
        return None
    except AuthorizationException:
        raise
    except Exception as exc:
        print(f"❌ Error with term '{term_name}': {exc}")
        return None
# ---------------------------------------------------------------------------
# Utility to strip default zeros
# ---------------------------------------------------------------------------
def _apply_qty_rate(detail: SalesItemLineDetail, qty: Optional[float], rate: Optional[float]):
    """Overwrite SDK's default values so blank CSV cols stay truly blank."""
    detail.Qty = None
    detail.UnitPrice = None
    detail.TaxInclusiveAmt = None
    detail.ServiceDate = None  # FIX: Prevent default ServiceDate from being sent
    if qty is not None:
        detail.Qty = qty
    if rate is not None:
        detail.UnitPrice = rate
# ---------------------------------------------------------------------------
# Core Invoice creation routine
# ---------------------------------------------------------------------------
def invoice_number_exists(client: QuickBooks, doc_number: str) -> bool:
    """Check if an invoice with the given DocNumber exists in QBO."""
    try:
        return bool(Invoice.filter(DocNumber=doc_number, qb=client))
    except AuthorizationException:
        raise
    except Exception as exc:
        print(f"❌ Error checking for existing invoice '{doc_number}': {exc}")
        return False
def create_quickbooks_invoice(
    client: QuickBooks, data: Dict[str, Any], inv_no: str, *,
    debug_json: bool, only_required: bool, auto_fill_qty_rate: bool
) -> bool:
    """Constructs and saves a single QuickBooks invoice."""
    if invoice_number_exists(client, inv_no):
        print(f"⚠️ Invoice number '{inv_no}' already exists. Skipping.")
        return False
    customer = find_or_create_customer(client, data["Customer"])
    if not customer: return False
    invoice = Invoice()
    invoice.CustomerRef = customer.to_ref()
    invoice.TxnDate = data["InvoiceDate"]
    invoice.DocNumber = inv_no
    if not only_required:
        invoice.DueDate = data["DueDate"]
        if data.get("CustomerMemo"):
            invoice.CustomerMemo = {"value": data["CustomerMemo"]}
        if data.get("Terms"):
            term = find_sales_term_by_name(client, data["Terms"])
            if term:
                invoice.SalesTermRef = term.to_ref()
    lines: List[SalesItemLine] = []
    for row in data["LineItems"]:
        item_name = row.get("Item")
        
        if item_name:
            # This is a standard line with a Product/Service item.
            line = SalesItemLine()
            detail = SalesItemLineDetail()
            item = find_or_create_item(client, item_name)
            if not item:
                print(f"🛑 Aborting invoice {inv_no} due to missing item '{item_name}'.")
                return False
            qty, rate, amount = row.get("Quantity"), row.get("Rate"), row["Amount"]
            if auto_fill_qty_rate:
                if qty is None and rate is None: qty, rate = 1, amount
                elif qty is None: qty = 1 if rate == 0 else round(amount / rate, 4)
                elif rate is None: rate = 0 if qty == 0 else round(amount / qty, 4)
            detail.ItemRef = item.to_ref()
            _apply_qty_rate(detail, qty, rate)
            if not only_required and row.get("Description"):
                line.Description = row["Description"]
            line.Amount = amount
            line.SalesItemLineDetail = detail
            lines.append(line)
        
        else:
            # This is a Description-Only line. ItemRef is omitted.
            description = row.get("Description")
            if not description:
                print("⚠️ Skipping line with no Item and no Description.")
                continue
            line = SalesItemLine()
            line.DetailType = 'DescriptionOnly'  # Set the correct DetailType
            line.Description = description
            line.Amount = row["Amount"]
            # CRITICAL: Do NOT attach a SalesItemLineDetail object.
            # The library will correctly serialize this as a DescriptionOnly line.
            lines.append(line)
    if not lines:
        print("⚠️ Invoice skipped – no valid line items detected.")
        return False
    invoice.Line = lines
    if debug_json:
        print("📤 JSON payload just before save:")
        print(json.dumps(json.loads(invoice.to_json()), indent=2))
    invoice.save(qb=client)
    print(f"✅ Created invoice {inv_no} (QBO Id: {invoice.Id})")
    return True
# ---------------------------------------------------------------------------
# Batch processing & CLI
# ---------------------------------------------------------------------------
def process_invoices(
    client: QuickBooks, invoices_data: Dict[str, Any], *,
    debug_json: bool, only_required: bool, auto_fill_qty_rate: bool
) -> None:
    """Processes all invoices, with token refresh and single retry logic."""
    total = len(invoices_data)
    success = 0
    print(f"\n📋 Processing {total} invoices…")
    for inv_no, data in invoices_data.items():
        print(f"\n🔄 Processing Invoice {inv_no}…")
        try:
            # First attempt
            if create_quickbooks_invoice(
                client, data, inv_no,
                debug_json=debug_json, only_required=only_required,
                auto_fill_qty_rate=auto_fill_qty_rate
            ):
                success += 1
        except AuthorizationException:
            print("🚨 QB Auth Exception 401: Token may have expired.")

            # Attempt to refresh the token
            if not refresh_access_token(client, CONFIG):
                print("🛑 Aborting script because token refresh failed.")
                break  # Exit the main loop

            # Reload tokens and rebuild the client so the retry uses the
            # freshest credentials saved to qb_tokens.json
            if not load_tokens():
                print("🛑 Failed to reload refreshed tokens.")
                break
            client = initialize_quickbooks_client()
            if not client:
                print("🛑 Failed to reinitialize QuickBooks client after token refresh.")
                break

            # If refresh was successful, retry the operation ONCE.
            print("✅ Token refreshed. Retrying the same invoice...")
            try:
                if create_quickbooks_invoice(
                    client, data, inv_no,
                    debug_json=debug_json, only_required=only_required,
                    auto_fill_qty_rate=auto_fill_qty_rate
                ):
                    success += 1
            except AuthorizationException:
                print(f"❌ CRITICAL: Auth failed again for invoice {inv_no} after token refresh.")
                print("   The company may have been disconnected or permissions revoked.")
                print("🛑 Aborting script.")
                break # Exit the main loop
            except Exception as exc:
                print(f"❌ Unexpected error on retry for invoice {inv_no}: {exc}")
        except Exception as exc:
            print(f"❌ An unexpected error occurred on invoice {inv_no}: {exc}")
    
    print(f"\n📊 Summary: {success}/{total} invoices processed.")
# ---------------------------------------------------------------------------
# Main entry-point
# ---------------------------------------------------------------------------
def main() -> None:
    """Main script execution function."""
    parser = argparse.ArgumentParser(description="Bulk-import invoices into QBO")
    parser.add_argument("--debug-json", action="store_true", help="Print JSON payload before save")
    parser.add_argument("--only-required", action="store_true", help="Send only mandatory fields")
    parser.add_argument("--auto-fill-qty-rate", action="store_true", help="Back-fill Qty/Rate when blank (legacy)")
    args = parser.parse_args()
    print("=== QuickBooks Invoice Import Script ===")
    if not validate_environment(): sys.exit(1)
    if not load_tokens() and not setup_oauth(): sys.exit(1)
    csv_file = "invoices.csv"
    if not os.path.exists(csv_file):
        print(f"❌ CSV file '{csv_file}' not found")
        sys.exit(1)
    print(f"📄 Loading invoices from {csv_file}…")
    invoices = read_invoices(csv_file)
    print(f"✅ Loaded {len(invoices)} invoices from CSV")
    client = initialize_quickbooks_client()
    if not client: sys.exit(1)
    process_invoices(
        client, invoices,
        debug_json=args.debug_json,
        only_required=args.only_required,
        auto_fill_qty_rate=args.auto_fill_qty_rate
    )
    print("\n🏁 Script completed!")
if __name__ == "__main__":
    main()
