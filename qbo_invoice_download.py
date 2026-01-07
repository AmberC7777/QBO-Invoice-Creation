#!/usr/bin/env python3
"""
QuickBooks Online – Invoice PDF Downloader
===========================================
This script reads a CSV file to download specified invoices from QuickBooks Online
as PDF files. It reuses authentication and configuration logic from the bulk
invoice import script.
CSV Format:
-----------
The script requires a `download.csv` file with two columns:
- InvoiceNo: The invoice number as it appears in QBO.
- FileName: The desired local name for the saved PDF file.
Example `download.csv`:
-----------------------
InvoiceNo,FileName
"1001","Invoice_1001.pdf"
"1002","Invoice_1002.pdf"
"""
from __future__ import annotations
import csv
import json
import os
import sys
from typing import Any, Dict, Optional
import requests
from dotenv import load_dotenv
from intuitlib.client import AuthClient
from quickbooks import QuickBooks
from quickbooks.objects import Invoice
from quickbooks.exceptions import AuthorizationException
from qb_auth import refresh_access_token, TOKEN_FILE

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
# OAuth Helpers (reused from import script)
# ---------------------------------------------------------------------------
def load_tokens() -> bool:
    """Load tokens from ``qb_tokens.json``."""
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
    """Guide user to set up OAuth tokens if missing."""
    print("=== OAuth Setup Required ===")
    print(
        "Run the OAuth playground (https://appcenter.intuit.com/playground) "
        "and save the resulting tokens to qb_tokens.json before continuing."
    )
    return False
# ---------------------------------------------------------------------------
# QuickBooks Client (reused from import script)
# ---------------------------------------------------------------------------
def initialize_quickbooks_client() -> Optional[QuickBooks]:
    """Initialize and return a QuickBooks client instance."""
    if not CONFIG["ACCESS_TOKEN"] or not CONFIG["REALM_ID"]:
        print("❌ Missing OAuth tokens. Please run the OAuth setup.")
        return None
    try:
        auth_client = AuthClient(
            client_id=str(CONFIG["CLIENT_ID"]),
            client_secret=str(CONFIG["CLIENT_SECRET"]),
            redirect_uri=str(CONFIG["REDIRECT_URI"]),
            environment="sandbox" if CONFIG["SANDBOX"] else "production",
            access_token=str(CONFIG["ACCESS_TOKEN"]),
            refresh_token=str(CONFIG["REFRESH_TOKEN"]),
        )
        # The token is automatically refreshed if expired
        CONFIG["ACCESS_TOKEN"] = auth_client.access_token
        return QuickBooks(
            auth_client=auth_client,
            refresh_token=str(CONFIG["REFRESH_TOKEN"]),
            company_id=str(CONFIG["REALM_ID"]),
            sandbox=bool(CONFIG["SANDBOX"]),
        )
    except Exception as e:
        print(f"❌ Failed to initialize QuickBooks client: {e}")
        return None
# ---------------------------------------------------------------------------
# Core Logic
# ---------------------------------------------------------------------------
def get_invoice_id(client: QuickBooks, invoice_no: str) -> Optional[str]:
    """Fetch the internal QBO ID for a given invoice number (DocNumber)."""
    try:
        invoices = Invoice.filter(DocNumber=invoice_no, qb=client)
        if invoices:
            return invoices[0].Id
        print(f"🔍 Invoice with number '{invoice_no}' not found in QBO.")
        return None
    except AuthorizationException:
        raise
    except Exception as e:
        print(f"❌ Error looking up invoice '{invoice_no}': {e}")
        return None
def download_invoice_pdf(client: QuickBooks, invoice_id: str) -> Optional[bytes]:
    """Download the PDF for a given invoice ID using a direct HTTP request."""
    try:
        base_url = "https://sandbox-quickbooks.api.intuit.com" if CONFIG["SANDBOX"] else "https://quickbooks.api.intuit.com"
        url = f"{base_url}/v3/company/{CONFIG['REALM_ID']}/invoice/{invoice_id}/pdf"

        headers = {
            "Authorization": f"Bearer {CONFIG['ACCESS_TOKEN']}",
            "Accept": "application/pdf",
        }

        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Raises an HTTPError for bad responses
        return response.content
    except requests.exceptions.HTTPError as e:
        if e.response is not None and e.response.status_code == 401:
            raise AuthorizationException("401 Unauthorized") from e
        print(f"❌ Failed to download PDF for invoice ID '{invoice_id}': {e}")
        return None
    except requests.exceptions.RequestException as e:
        print(f"❌ Failed to download PDF for invoice ID '{invoice_id}': {e}")
        return None
def get_unique_filename(file_path: str) -> str:
    """
    Check if a file exists. If so, append a counter to find a unique name.
    Example: "file.pdf" -> "file(1).pdf" -> "file(2).pdf"
    """
    if not os.path.exists(file_path):
        return file_path
    base, ext = os.path.splitext(file_path)
    counter = 1
    while True:
        new_path = f"{base}({counter}){ext}"
        if not os.path.exists(new_path):
            return new_path
        counter += 1
def process_invoices(client: QuickBooks, csv_path: str):
    """Read the CSV and process each invoice for download."""
    print(f"🚀 Starting invoice download process from '{csv_path}'...")

    def refresh_and_reinitialize(qb_client: QuickBooks) -> Optional[QuickBooks]:
        """Refresh tokens and rebuild the QuickBooks client."""
        if not refresh_access_token(qb_client, CONFIG):
            print("🛑 Aborting process because token refresh failed.")
            return None
        if not load_tokens():
            print("🛑 Failed to reload refreshed tokens.")
            return None
        new_client = initialize_quickbooks_client()
        if not new_client:
            print("🛑 Failed to reinitialize QuickBooks client after token refresh.")
            return None
        return new_client

    try:
        with open(csv_path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                invoice_no = row.get("InvoiceNo", "").strip()
                file_name = row.get("FileName", "").strip()
                if not invoice_no or not file_name:
                    print("⚠️ Skipping row with missing InvoiceNo or FileName.")
                    continue

                print(f"\nProcessing InvoiceNo: {invoice_no}")

                try:
                    invoice_id = get_invoice_id(client, invoice_no)
                except AuthorizationException:
                    print("🚨 QB Auth Exception 401: Token may have expired.")
                    refreshed_client = refresh_and_reinitialize(client)
                    if not refreshed_client:
                        break
                    client = refreshed_client
                    print("✅ Token refreshed. Retrying the same invoice...")
                    try:
                        invoice_id = get_invoice_id(client, invoice_no)
                    except AuthorizationException:
                        print("❌ CRITICAL: Auth failed again after token refresh.")
                        break
                if not invoice_id:
                    continue  # Error already logged by get_invoice_id

                try:
                    pdf_content = download_invoice_pdf(client, invoice_id)
                except AuthorizationException:
                    print("🚨 QB Auth Exception 401: Token may have expired.")
                    refreshed_client = refresh_and_reinitialize(client)
                    if not refreshed_client:
                        break
                    client = refreshed_client
                    print("✅ Token refreshed. Retrying the same invoice...")
                    try:
                        pdf_content = download_invoice_pdf(client, invoice_id)
                    except AuthorizationException:
                        print("❌ CRITICAL: Auth failed again after token refresh.")
                        break
                if not pdf_content:
                    continue  # Error already logged by download_invoice_pdf

                # Ensure the 'invoices' directory exists
                output_dir = "invoices"
                os.makedirs(output_dir, exist_ok=True)
                
                full_path = os.path.join(output_dir, file_name)
                unique_path = get_unique_filename(full_path)
                try:
                    with open(unique_path, "wb") as pdf_file:
                        pdf_file.write(pdf_content)
                    print(f"✅ Successfully downloaded and saved to '{unique_path}'")
                except IOError as e:
                    print(f"❌ Failed to save PDF to '{unique_path}': {e}")
    except FileNotFoundError:
        print(f"❌ Error: The file '{csv_path}' was not found.")
    except Exception as e:
        print(f"❌ An unexpected error occurred while processing the CSV: {e}")
def main():
    """Main execution function."""
    if not validate_environment():
        sys.exit(1)
    if not load_tokens():
        setup_oauth()
        sys.exit(1)
    client = initialize_quickbooks_client()
    if not client:
        sys.exit(1)
    print("✅ QuickBooks client initialized successfully.")
    
    process_invoices(client, "download.csv")
    print("\n🎉 Invoice download process complete.")

if __name__ == "__main__":
    main()
