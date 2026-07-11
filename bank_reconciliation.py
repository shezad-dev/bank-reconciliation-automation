#!/usr/bin/env python3
"""
Bank Statement Reconciliation

A Python script that automatically reconciles bank transactions against invoices.
Matches transactions by invoice number and amount, generates a PDF report,
and emails it to the finance team.

Features:
- Reads data from Google Sheets
- Extracts invoice numbers from transaction descriptions
- Matches bank transactions to invoices
- Flags amount mismatches
- Identifies missing invoice references
- Generates professional PDF report with tables
- Emails report as PDF attachment

Author: Shezad Dev
GitHub: github.com/shezad-dev
"""

# ============ IMPORT LIBRARIES ============
import urllib.request
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import os
import re
from fpdf import FPDF

# ============ CONFIGURATION ============
# Replace these with your own credentials
BANK_SHEET_ID = "YOUR_BANK_SHEET_ID_HERE"
INVOICE_SHEET_ID = "YOUR_INVOICE_SHEET_ID_HERE"

GMAIL_USER = "your_email@gmail.com"
GMAIL_PASSWORD = "your_app_password"
ALERT_EMAIL = "manager@company.com"

# ============ FUNCTION: READ GOOGLE SHEET ============
def read_sheet(sheet_id, sheet_name):
    """
    Reads a public Google Sheet tab as CSV.
    Returns list of rows, each row is a list of columns.
    """
    try:
        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
        headers = {'User-Agent': 'Mozilla/5.0'}
        req = urllib.request.Request(url, headers=headers)
        response = urllib.request.urlopen(req)
        data = response.read().decode('utf-8')
        
        lines = data.strip().split('\n')
        result = []
        
        for line in lines:
            parts = []
            current = ''
            in_quotes = False
            
            for char in line:
                if char == '"' and not in_quotes:
                    in_quotes = True
                elif char == '"' and in_quotes:
                    in_quotes = False
                elif char == ',' and not in_quotes:
                    parts.append(current.strip())
                    current = ''
                else:
                    current += char
            
            parts.append(current.strip())
            result.append(parts)
        
        return result
        
    except Exception as e:
        print(f"❌ Error reading sheet: {e}")
        return None

# ============ FUNCTION: EXTRACT INVOICE NUMBER ============
def extract_invoice_number(description):
    """
    Extracts invoice number from transaction description.
    Returns invoice number string, or None.
    """
    patterns = [
        r'INV-\d+',
        r'INV\d+',
        r'INVOICE\s*#?\s*\d+',
        r'INVOICE\s*-\s*\d+',
        r'Invoice\s*#?\s*\d+',
        r'INV\s*#?\s*\d+'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, description, re.IGNORECASE)
        if match:
            return match.group(0).upper()
    
    return None

# ============ FUNCTION: CREATE PDF REPORT ============
def create_pdf_report(report_data, filename):
    """
    Creates a professional PDF report with tables.
    """
    try:
        pdf = FPDF()
        pdf.add_page()
        
        # Title
        pdf.set_font("Helvetica", "B", 18)
        pdf.cell(0, 12, "Bank Reconciliation Report", ln=True, align="C")
        pdf.ln(2)
        
        # Date
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 8, f"Generated: {report_data['timestamp']}", ln=True, align="C")
        pdf.ln(8)
        
        # Divider
        pdf.line(20, pdf.get_y(), 190, pdf.get_y())
        pdf.ln(6)
        
        # Summary
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 10, "SUMMARY", ln=True)
        pdf.set_font("Helvetica", "", 11)
        
        summary_items = [
            ("Total Bank Transactions: ", report_data['total']),
            ("Matched: ", report_data['matched_count']),
            ("Amount Mismatch: ", report_data['amount_mismatch_count']),
            ("Invoice Not Found: ", report_data['inv_not_found_count']),
            ("No INV Reference: ", report_data['no_ref_count'])
        ]
        
        for label, value in summary_items:
            pdf.set_font("Helvetica", "B", 11)
            pdf.cell(60, 8, label, border=0)
            pdf.set_font("Helvetica", "", 11)
            pdf.cell(0, 8, str(value), ln=True)
        
        pdf.ln(4)
        pdf.line(20, pdf.get_y(), 190, pdf.get_y())
        pdf.ln(6)
        
        # Amount Mismatch Table
        if report_data['amount_mismatch']:
            pdf.set_font("Helvetica", "B", 12)
            pdf.cell(0, 10, "Amount Mismatch Details:", ln=True)
            pdf.ln(2)
            
            pdf.set_font("Helvetica", "B", 9)
            pdf.set_fill_color(200, 200, 200)
            pdf.cell(25, 8, "Date", border=1, fill=True)
            pdf.cell(50, 8, "Description", border=1, fill=True)
            pdf.cell(30, 8, "Bank", border=1, fill=True)
            pdf.cell(30, 8, "Invoice", border=1, fill=True)
            pdf.cell(25, 8, "Diff", border=1, fill=True, ln=True)
            
            pdf.set_font("Helvetica", "", 8)
            for item in report_data['amount_mismatch'][:20]:
                diff = abs(abs(item['amount']) - item['inv_amount'])
                pdf.cell(25, 6, item['date'], border=1)
                pdf.cell(50, 6, item['desc'][:20], border=1)
                pdf.cell(30, 6, f"£{abs(item['amount']):.2f}", border=1)
                pdf.cell(30, 6, f"£{item['inv_amount']:.2f}", border=1)
                pdf.cell(25, 6, f"£{diff:.2f}", border=1, ln=True)
            
            if len(report_data['amount_mismatch']) > 20:
                pdf.cell(0, 6, f"... and {len(report_data['amount_mismatch'])-20} more", ln=True)
            pdf.ln(4)
        
        # Invoice Not Found Table
        if report_data['invoice_not_found']:
            pdf.set_font("Helvetica", "B", 12)
            pdf.cell(0, 10, "Invoice Not Found:", ln=True)
            pdf.ln(2)
            
            pdf.set_font("Helvetica", "B", 9)
            pdf.set_fill_color(200, 200, 200)
            pdf.cell(25, 8, "Date", border=1, fill=True)
            pdf.cell(50, 8, "Description", border=1, fill=True)
            pdf.cell(30, 8, "Amount", border=1, fill=True)
            pdf.cell(40, 8, "Invoice #", border=1, fill=True, ln=True)
            
            pdf.set_font("Helvetica", "", 8)
            for item in report_data['invoice_not_found'][:20]:
                pdf.cell(25, 6, item['date'], border=1)
                pdf.cell(50, 6, item['desc'][:20], border=1)
                pdf.cell(30, 6, f"£{abs(item['amount']):.2f}", border=1)
                pdf.cell(40, 6, item['inv_num'], border=1, ln=True)
            
            if len(report_data['invoice_not_found']) > 20:
                pdf.cell(0, 6, f"... and {len(report_data['invoice_not_found'])-20} more", ln=True)
            pdf.ln(4)
        
        # No Invoice Reference Table
        if report_data['no_invoice_ref']:
            pdf.set_font("Helvetica", "B", 12)
            pdf.cell(0, 10, "No Invoice Reference:", ln=True)
            pdf.ln(2)
            
            pdf.set_font("Helvetica", "B", 9)
            pdf.set_fill_color(200, 200, 200)
            pdf.cell(25, 8, "Date", border=1, fill=True)
            pdf.cell(60, 8, "Description", border=1, fill=True)
            pdf.cell(30, 8, "Amount", border=1, fill=True, ln=True)
            
            pdf.set_font("Helvetica", "", 8)
            for item in report_data['no_invoice_ref'][:20]:
                pdf.cell(25, 6, item['date'], border=1)
                pdf.cell(60, 6, item['desc'][:25], border=1)
                pdf.cell(30, 6, f"£{abs(item['amount']):.2f}", border=1, ln=True)
            
            if len(report_data['no_invoice_ref']) > 20:
                pdf.cell(0, 6, f"... and {len(report_data['no_invoice_ref'])-20} more", ln=True)
            pdf.ln(4)
        
        # Divider
        pdf.line(20, pdf.get_y(), 190, pdf.get_y())
        pdf.ln(6)
        
        # Matched Count
        if report_data['matched_count'] > 0:
            pdf.set_font("Helvetica", "B", 12)
            pdf.set_text_color(0, 150, 0)
            pdf.cell(0, 10, f"Matched Transactions: {report_data['matched_count']} matched successfully.", ln=True)
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Helvetica", "", 10)
            pdf.cell(0, 6, "No action needed for these.", ln=True)
        
        pdf.ln(4)
        
        # Footer
        pdf.set_y(-20)
        pdf.set_font("Helvetica", "I", 8)
        pdf.line(20, pdf.get_y(), 190, pdf.get_y())
        pdf.ln(2)
        pdf.cell(0, 8, "This is an automated reconciliation report.", align="C")
        
        pdf.output(filename)
        return True
        
    except Exception as e:
        print(f"❌ PDF creation error: {e}")
        return False

# ============ FUNCTION: SEND EMAIL WITH PDF ============
def send_email_with_pdf(subject, body, pdf_path):
    """
    Sends an email with PDF attachment.
    """
    try:
        msg = MIMEMultipart()
        msg['From'] = GMAIL_USER
        msg['To'] = ALERT_EMAIL
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        
        with open(pdf_path, 'rb') as f:
            part = MIMEBase('application', 'pdf')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(pdf_path)}')
            msg.attach(part)
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"❌ Email failed: {e}")
        return False

# ============ MAIN RECONCILIATION FUNCTION ============
def run_reconciliation():
    """
    Main reconciliation function.
    """
    print("\n" + "="*70)
    print("  BANK RECONCILIATION WITH PDF REPORT")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*70 + "\n")
    
    # Step 1: Read data
    print("📥 Reading Google Sheets...")
    bank_data = read_sheet(BANK_SHEET_ID, "Bank_Statement")
    invoice_data = read_sheet(INVOICE_SHEET_ID, "Invoices")
    
    if not bank_data or not invoice_data:
        print("❌ Missing data.")
        return
    
    bank_rows = bank_data[1:]
    invoice_rows = invoice_data[1:]
    
    print(f"   ✅ Bank Transactions: {len(bank_rows)}")
    print(f"   ✅ Invoices: {len(invoice_rows)}\n")
    
    # Step 2: Build invoice lookup
    print("🔍 Building invoice lookup...")
    invoice_map = {}
    
    for row in invoice_rows:
        if len(row) < 3:
            continue
        
        inv_num = row[0].strip().upper() if row[0] else ""
        if not inv_num:
            continue
        
        try:
            amount = float(row[2]) if row[2] else 0
        except:
            amount = 0
        
        invoice_map[inv_num] = {
            'amount': amount,
            'client': row[1].strip() if len(row) > 1 and row[1] else "",
            'date': row[3].strip() if len(row) > 3 and row[3] else "",
            'status': row[4].strip() if len(row) > 4 and row[4] else "",
            'description': row[5].strip() if len(row) > 5 and row[5] else ""
        }
    
    print(f"   ✅ {len(invoice_map)} invoices loaded\n")
    
    # Step 3: Match transactions
    print("🔍 Matching transactions...\n")
    
    matched = []
    amount_mismatch = []
    invoice_not_found = []
    no_invoice_ref = []
    
    matched_count = 0
    amount_mismatch_count = 0
    inv_not_found_count = 0
    no_ref_count = 0
    
    print("-"*90)
    print(f"{'Date':<12} {'Description':<35} {'Amount':<12} {'Status':<25}")
    print("-"*90)
    
    for row in bank_rows:
        if len(row) < 3:
            continue
        
        date = row[0].strip() if row[0] else ""
        desc = row[1].strip() if row[1] else ""
        try:
            amount = float(row[2]) if row[2] else 0
        except:
            amount = 0
        
        inv_num = extract_invoice_number(desc)
        
        if inv_num:
            if inv_num in invoice_map:
                inv_data = invoice_map[inv_num]
                inv_amount = inv_data['amount']
                
                if abs(abs(amount) - inv_amount) < 0.01:
                    status = "✅ MATCHED"
                    matched_count += 1
                    matched.append({
                        'date': date,
                        'desc': desc,
                        'amount': amount,
                        'inv_num': inv_num,
                        'inv_amount': inv_amount
                    })
                else:
                    status = "❌ AMOUNT MISMATCH"
                    amount_mismatch_count += 1
                    amount_mismatch.append({
                        'date': date,
                        'desc': desc,
                        'amount': amount,
                        'inv_num': inv_num,
                        'inv_amount': inv_amount
                    })
            else:
                status = "❌ INVOICE NOT FOUND"
                inv_not_found_count += 1
                invoice_not_found.append({
                    'date': date,
                    'desc': desc,
                    'amount': amount,
                    'inv_num': inv_num
                })
        else:
            status = "⚠️ NO INV REF"
            no_ref_count += 1
            no_invoice_ref.append({
                'date': date,
                'desc': desc,
                'amount': amount
            })
        
        if len(matched) + len(amount_mismatch) + len(invoice_not_found) + len(no_invoice_ref) <= 20:
            print(f"{date:<12} {desc[:33]:<35} {amount:<12.2f} {status:<25}")
    
    print("-"*90)
    
    total = len(bank_rows)
    
    print("\n" + "="*70)
    print("📊 SUMMARY")
    print("="*70)
    print(f"Total Bank Transactions:  {total}")
    print(f"✅ Matched:                {matched_count}")
    print(f"❌ Amount Mismatch:        {amount_mismatch_count}")
    print(f"❌ Invoice Not Found:      {inv_not_found_count}")
    print(f"⚠️ No INV Reference:        {no_ref_count}")
    print("="*70 + "\n")
    
    # Step 4: Generate PDF
    print("📄 Generating PDF report...")
    
    report_data = {
        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'total': total,
        'matched_count': matched_count,
        'amount_mismatch_count': amount_mismatch_count,
        'inv_not_found_count': inv_not_found_count,
        'no_ref_count': no_ref_count,
        'amount_mismatch': amount_mismatch,
        'invoice_not_found': invoice_not_found,
        'no_invoice_ref': no_invoice_ref
    }
    
    pdf_filename = f"reconciliation_report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
    pdf_created = create_pdf_report(report_data, pdf_filename)
    
    if not pdf_created:
        print("❌ PDF creation failed.")
        return
    
    print(f"✅ PDF created: {pdf_filename}")
    
    # Step 5: Send email
    print("📧 Sending email with PDF attachment...")
    
    body = f"""
BANK RECONCILIATION REPORT
======================================

📅 Report Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

📊 SUMMARY
======================================

Total Bank Transactions:  {total}
✅ Matched:                {matched_count}
❌ Amount Mismatch:        {amount_mismatch_count}
❌ Invoice Not Found:      {inv_not_found_count}
⚠️ No INV Reference:       {no_ref_count}

======================================

📎 PDF report attached: {os.path.basename(pdf_filename)}
"""
    
    send_email_with_pdf(
        f"Bank Reconciliation Report - {matched_count} matched, {amount_mismatch_count + inv_not_found_count + no_ref_count} issues",
        body,
        pdf_filename
    )
    
    print("✅ Email sent with PDF attachment!")
    print("\n✅ Reconciliation complete.")

# ============ RUN THE SCRIPT ============
if __name__ == "__main__":
    run_reconciliation()
