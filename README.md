# Bank Reconciliation Automation

Automated bank statement reconciliation that matches transactions to invoices, flags mismatches, and generates PDF reports.

## Features

- Reads data from Google Sheets
- Extracts invoice numbers from transaction descriptions
- Matches bank transactions to invoices by amount
- Flags amount mismatches
- Identifies transactions with no invoice reference
- Generates professional PDF report with tables
- Emails report as PDF attachment to finance team

## How It Works

```

Google Sheets (Bank Statement + Invoices)
↓
Python Script (Reads data, matches transactions)
↓
PDF Report (Professional formatted tables)
↓
Email (PDF attached to finance team)

```

## Requirements

- Python 3.6+
- fpdf2 library

## Installation

1. Clone this repository:
```bash
git clone https://github.com/shezad-dev/bank-reconciliation-automation.git
cd bank-reconciliation-automation
```

2. Install dependencies:

```bash
pip3 install fpdf2
```

Configuration

1. Create a Google Sheet with two tabs:
   · Bank_Statement - Columns: Date, Description, Amount, Type, Balance
   · Invoices - Columns: Invoice_Number, Client, Amount, Due_Date, Status, Description
2. Make the sheet public (Share → Anyone with link can view)
3. Get your Sheet ID from the URL:
   ```
   https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit
   ```
4. Create config.py in the same folder:

```python
# config.py
GMAIL_USER = "your_email@gmail.com"
GMAIL_PASSWORD = "your_app_password"  # Get from Google App Passwords
ALERT_EMAIL = "manager@company.com"
```

5. Update the script with your Sheet IDs:

```python
BANK_SHEET_ID = "YOUR_BANK_SHEET_ID_HERE"
INVOICE_SHEET_ID = "YOUR_INVOICE_SHEET_ID_HERE"
```

Google Sheet Structure

Bank_Statement Tab

Date Description Amount Type Balance
2026-07-01 PAYMENT - INV-001 -1200.00 Debit 8800.00
2026-07-02 POS purchase - Amazon -150.00 Debit 8650.00

Invoices Tab

Invoice_Number Client Amount Due_Date Status Description
INV-001 Acme Corp 1200.00 2026-07-15 Paid Website Development

How to Get App Password

1. Go to Google Account → Security
2. Turn on 2-Step Verification
3. Go to App Passwords
4. Select "Mail" and "Android"
5. Copy the 16-character password

Running the Script

```bash
python3 bank_reconciliation.py
```

Output

· On Screen: Summary of matched and unmatched transactions
· Email: PDF report with all details
· Local: PDF saved to /storage/emulated/0/

PDF Report Example

```
======================================================================
                  Bank Reconciliation Report
                  Generated: 2026-07-11 22:29:15
======================================================================

SUMMARY
----------------------------------------
Total Bank Transactions:  400
Matched:                  372
Amount Mismatch:          6
Invoice Not Found:        0
No INV Reference:         22

======================================================================

Amount Mismatch Details:
┌──────────┬─────────────────────┬──────────┬──────────┬────────┐
│ Date     │ Description         │ Bank     │ Invoice  │ Diff   │
├──────────┼─────────────────────┼──────────┼──────────┼────────┤
│2026-07-03│ PAYMENT - INV-038   │ £1085.17 │ £1090.91 │ £5.74  │
└──────────┴─────────────────────┴──────────┴──────────┴────────┘

======================================================================
              This is an automated reconciliation report.
```

Technologies Used

· Python 3
· Google Sheets API (CSV export)
· Gmail SMTP
· fpdf2 (PDF generation)

Author

Shezad Dev

· GitHub: shezad-dev

License

MIT License - Free for commercial and personal use.
