# QuickBooks Online Invoice Utilities

Python scripts for bulk importing invoices into QuickBooks Online and downloading invoice PDFs.

## Features

- Import invoices from `invoices.csv` with automatic customer and item creation.
- Token handling with OAuth 2.0, including automatic access token refresh on 401 errors.
- Correctly creates description-only line items and prevents unwanted default `ServiceDate`.
- Set payment terms by name via the `Terms` column in the CSV.
- Download invoice PDFs listed in `download.csv` and save them to the `invoices/` folder with unique filenames.
- Supports sandbox and production environments via configuration.

## Setup

### 1. Install Dependencies
```bash
pip install python-dotenv intuitlib quickbooks requests
```

### 2. Configure Environment Variables
Copy the example file and edit with your credentials:
```bash
cp .env.example .env
```

### 3. Get QuickBooks API Credentials
1. Go to [Intuit Developer](https://developer.intuit.com/)
2. Create a new app or use an existing one
3. Note your Client ID and Client Secret
4. Add the redirect URI to your app settings

### 4. Set Up OAuth
1. Run an import or download script for the first time
2. Follow the printed instructions to use the OAuth playground
3. Save the resulting tokens to `qb_tokens.json`

## Usage

### Import Invoices
```bash
python qbo_invoice_import.py [--debug-json] [--only-required] [--auto-fill-qty-rate]
```

### Download Invoice PDFs
Prepare `download.csv` with two columns: `InvoiceNo` and `FileName`, then run:
```bash
python qbo_invoice_download.py
```

PDF files are saved in the `invoices/` directory, which is created if missing. Existing filenames are given a numeric suffix to avoid overwriting.

## CSV Format for Import
Required columns in `invoices.csv`:
- `InvoiceNo`: Invoice number
- `Customer`: Customer name
- `InvoiceDate`: MM/DD/YYYY
- `DueDate`: MM/DD/YYYY
- `CustomerMemo`: Optional memo
- `Terms`: Optional payment term name
- `Item(Product/Service)`: Item name (leave blank for description-only lines)
- `ItemDescription`: Optional description (put "Subtotal:" for in-line subtotal (a description-only line))
- `ItemQuantity`: Optional quantity
- `ItemRate`: Optional rate
- `ItemAmount`: Line amount

If `Terms` does not match a term in QuickBooks, the invoice is created without a payment term.

## Security
- Credentials are stored in `.env` (not committed to git)
- OAuth tokens are kept in `qb_tokens.json`
- Avoid committing sensitive files

## Troubleshooting

### Missing Environment Variables
1. Ensure `.env` exists and contains required variables
2. Verify `CLIENT_ID` and `CLIENT_SECRET`
3. Confirm the file is in the project root

### OAuth Problems
1. Check that the redirect URI matches your app settings
2. Ensure the `SANDBOX` flag matches your environment
3. Re-run OAuth setup if tokens expire

## File Structure
```
├── qbo_invoice_import.py     # Bulk invoice importer
├── qbo_invoice_download.py   # Invoice PDF downloader
├── qb_auth.py                # OAuth helper utilities
├── .env.example              # Sample environment file
└── README.md                 # Project documentation
```