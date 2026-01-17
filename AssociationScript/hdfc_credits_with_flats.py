#!/usr/bin/env python3
"""
Enhanced HDFC Credits Converter with Flat Number Detection
Extract credits and identify flat numbers (A001-A320, B001-B312, C001-C318)
"""

import sys
import os
from pathlib import Path
import re
import pandas as pd
from datetime import datetime

def extract_pdf_text():
    """Extract text from HDFC bank statement PDF"""
    print("📄 STEP 1: Extracting PDF Text")
    print("=" * 35)
    
    try:
        import pdfplumber
        
        input_dir = Path("input")
        pdf_files = list(input_dir.glob("*.pdf"))
        
        if not pdf_files:
            print("❌ No PDF files found")
            return None
        
        pdf_file = pdf_files[0]
        print(f"📂 Processing: {pdf_file.name}")
        
        full_text = ""
        with pdfplumber.open(pdf_file) as pdf:
            print(f"📄 Extracting from {len(pdf.pages)} pages...")
            
            for page_num, page in enumerate(pdf.pages, 1):
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text + "\n"
                if page_num % 5 == 0:
                    print(f"   ✅ Processed {page_num} pages")
        
        print(f"✅ Extraction complete")
        return full_text
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

def extract_flat_number(description):
    """Extract flat number from description using the specified patterns"""
    
    # Define the flat number patterns (removed word boundaries for embedded matches)
    flat_patterns = [
        # A001 to A320
        r'A(?:0(?:0[1-9]|[1-9][0-9])|[12][0-9][0-9]|3(?:[01][0-9]|20))',
        # B001 to B312  
        r'B(?:0(?:0[1-9]|[1-9][0-9])|[12][0-9][0-9]|3(?:0[0-9]|1[0-2]))',
        # C001 to C318
        r'C(?:0(?:0[1-9]|[1-9][0-9])|[12][0-9][0-9]|3(?:0[0-9]|1[0-8]))'
    ]
    
    # Search for flat numbers in the description
    for pattern in flat_patterns:
        match = re.search(pattern, description.upper())
        if match:
            return match.group()
    
    return ""

def extract_credits_with_flat_numbers(text):
    """Extract credit transactions and identify flat numbers"""
    print("\n💰 STEP 2: Extracting Credits with Flat Number Detection")
    print("=" * 60)
    
    lines = text.split('\n')
    lines = [line.strip() for line in lines if line.strip()]
    
    # Find all transaction lines
    transaction_lines = []
    for line in lines:
        if re.match(r'^\d{2}/\d{2}/\d{2}\s', line):
            transaction_lines.append(line)
    
    print(f"📊 Found {len(transaction_lines)} total transactions")
    
    credits = []
    
    # Process all transactions and compare balances
    for i, line in enumerate(transaction_lines):
        # Extract basic info
        date_match = re.match(r'^(\d{2}/\d{2}/\d{2})', line)
        if not date_match:
            continue
            
        date_str = date_match.group(1)
        amounts = re.findall(r'[\d,]+\.\d{2}', line)
        
        if len(amounts) < 2:
            continue
            
        transaction_amount = float(amounts[-2].replace(',', ''))
        current_balance = float(amounts[-1].replace(',', ''))
        
        # Compare with previous balance to determine if this is credit or debit
        is_credit = False
        balance_change = 0
        
        if i > 0:
            prev_line = transaction_lines[i-1]
            prev_amounts = re.findall(r'[\d,]+\.\d{2}', prev_line)
            if prev_amounts:
                prev_balance = float(prev_amounts[-1].replace(',', ''))
                balance_change = current_balance - prev_balance
                
                # If balance increased, it's a credit
                if balance_change > 0:
                    is_credit = True
        
        # For the first transaction, use pattern-based detection as fallback
        if i == 0:
            line_upper = line.upper()
            is_credit = any(indicator in line_upper for indicator in 
                          ['NEFTCR', 'RTGSCR', 'DEPOSIT', 'SALARY', 'INTEREST', 'CREDIT', 'CR-'])
        
        if is_credit:
            # Extract description
            description = line
            description = re.sub(r'^\d{2}/\d{2}/\d{2}\s+', '', description)
            for amt in reversed(amounts):
                description = description.replace(amt, '', 1)
            description = re.sub(r'\s+', ' ', description).strip()
            description = description.replace('01/12/25', '').strip()
            
            # Extract flat number from description
            flat_number = extract_flat_number(description)
            
            credits.append({
                'Date': date_str,
                'Narration': description,
                'Credit Amount (₹)': abs(balance_change) if balance_change != 0 else transaction_amount,
                'Account Balance (₹)': current_balance,
                'Flat Number': flat_number
            })
    
    print(f"💳 Found {len(credits)} credit transactions")
    
    # Show statistics about flat number detection
    credits_with_flats = [c for c in credits if c['Flat Number']]
    print(f"🏠 Found {len(credits_with_flats)} transactions with flat numbers")
    
    if credits_with_flats:
        print(f"\n📋 Sample transactions with flat numbers:")
        for i, credit in enumerate(credits_with_flats[:5]):
            print(f"{i+1}. {credit['Date']} - {credit['Flat Number']} - ₹{credit['Credit Amount (₹)']:,.2f}")
            print(f"   Description: {credit['Narration'][:60]}...")
    
    # Show flat number distribution
    flat_numbers = [c['Flat Number'] for c in credits if c['Flat Number']]
    if flat_numbers:
        from collections import Counter
        flat_distribution = Counter(flat_numbers)
        print(f"\n🏠 FLAT NUMBER DISTRIBUTION:")
        for flat_num, count in sorted(flat_distribution.items()):
            print(f"   {flat_num}: {count} transactions")
    
    return credits

def create_flat_credits_excel(credits):
    """Create Excel file with flat number extraction"""
    print(f"\n📊 STEP 3: Creating Excel with Flat Numbers")
    print("=" * 45)
    
    if not credits:
        print("❌ No credits found")
        return None
    
    df = pd.DataFrame(credits)
    
    # Sort by date
    df['Date_Sort'] = pd.to_datetime(df['Date'], format='%d/%m/%y')
    df = df.sort_values('Date_Sort').drop('Date_Sort', axis=1)
    
    # Create output file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = Path("output") / f"HDFC_Credits_with_Flats_{timestamp}.xlsx"
    output_file.parent.mkdir(exist_ok=True)
    
    # Create Excel with formatting
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Main sheet
        df.to_excel(writer, sheet_name='Credit Transactions', index=False)
        
        # Create separate sheet for transactions with flat numbers
        flats_df = df[df['Flat Number'] != ''].copy()
        if not flats_df.empty:
            flats_df.to_excel(writer, sheet_name='Transactions with Flats', index=False)
        
        # Format main sheet
        workbook = writer.book
        worksheet = writer.sheets['Credit Transactions']
        
        # Header format
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'top',
            'fg_color': '#4CAF50', 'font_color': 'white', 'border': 1, 'font_size': 11
        })
        
        # Rupee format
        rupee_format = workbook.add_format({
            'num_format': '₹#,##,##0.00', 'align': 'right', 'font_color': '#2E7D32'
        })
        
        # Flat number format (highlight if present)
        flat_format = workbook.add_format({
            'bg_color': '#FFF9C4', 'align': 'center', 'bold': True, 'font_color': '#F57F17'
        })
        
        # Apply header formatting
        for col_num, value in enumerate(df.columns):
            worksheet.write(0, col_num, value, header_format)
        
        # Set column widths and formats
        worksheet.set_column('A:A', 12)      # Date
        worksheet.set_column('B:B', 50)      # Narration
        worksheet.set_column('C:C', 18, rupee_format)  # Credit Amount
        worksheet.set_column('D:D', 20, rupee_format)  # Account Balance
        worksheet.set_column('E:E', 12, flat_format)   # Flat Number
        
        # Add conditional formatting for flat numbers
        worksheet.conditional_format('E2:E1000', {
            'type': 'cell',
            'criteria': '!=',
            'value': '""',
            'format': flat_format
        })
        
        # Summary section
        summary_row = len(df) + 3
        total_credits = df['Credit Amount (₹)'].sum()
        credits_with_flats = len(df[df['Flat Number'] != ''])
        
        summary_header = workbook.add_format({
            'bold': True, 'font_size': 12, 'fg_color': '#E8F5E8'
        })
        
        worksheet.write(summary_row, 0, "SUMMARY:", summary_header)
        worksheet.write(summary_row + 1, 0, "Total Credit Transactions:")
        worksheet.write(summary_row + 1, 1, len(credits))
        
        worksheet.write(summary_row + 2, 0, "Total Credits Amount:")
        worksheet.write(summary_row + 2, 1, total_credits, rupee_format)
        
        worksheet.write(summary_row + 3, 0, "Transactions with Flat Numbers:")
        worksheet.write(summary_row + 3, 1, credits_with_flats)
        
        # Flat number breakdown
        if credits_with_flats > 0:
            worksheet.write(summary_row + 5, 0, "FLAT NUMBER BREAKDOWN:", summary_header)
            
            flat_summary = df[df['Flat Number'] != ''].groupby('Flat Number')['Credit Amount (₹)'].agg(['count', 'sum'])
            
            for i, (flat_num, data) in enumerate(flat_summary.iterrows()):
                worksheet.write(summary_row + 6 + i, 0, f"{flat_num}:")
                worksheet.write(summary_row + 6 + i, 1, f"{data['count']} transactions")
                worksheet.write(summary_row + 6 + i, 2, data['sum'], rupee_format)
        
        # Format the flats sheet if it exists
        if not flats_df.empty:
            flats_worksheet = writer.sheets['Transactions with Flats']
            
            # Apply same formatting to flats sheet
            for col_num, value in enumerate(flats_df.columns):
                flats_worksheet.write(0, col_num, value, header_format)
            
            flats_worksheet.set_column('A:A', 12)
            flats_worksheet.set_column('B:B', 50)
            flats_worksheet.set_column('C:C', 18, rupee_format)
            flats_worksheet.set_column('D:D', 20, rupee_format)
            flats_worksheet.set_column('E:E', 12, flat_format)
    
    print(f"✅ Excel created: {output_file}")
    print(f"💰 Total Credits: ₹{total_credits:,.2f}")
    print(f"📊 Total Transactions: {len(credits)}")
    print(f"🏠 Transactions with Flat Numbers: {credits_with_flats}")
    
    if credits_with_flats > 0:
        flat_summary = df[df['Flat Number'] != ''].groupby('Flat Number')['Credit Amount (₹)'].agg(['count', 'sum'])
        print(f"\n🏠 FLAT NUMBER BREAKDOWN:")
        for flat_num, data in flat_summary.iterrows():
            print(f"   {flat_num}: {data['count']} transactions = ₹{data['sum']:,.2f}")
    
    return output_file

def main():
    """Main function"""
    print("🏦 HDFC CREDITS CONVERTER WITH FLAT NUMBERS")
    print("=" * 50)
    print("🎯 Extract credits and detect flat numbers (A001-A320, B001-B312, C001-C318)")
    print()
    
    # Extract text
    text = extract_pdf_text()
    if not text:
        return
    
    # Extract credits with flat number detection
    credits = extract_credits_with_flat_numbers(text)
    if not credits:
        print("❌ No credits found")
        return
    
    # Create Excel
    excel_file = create_flat_credits_excel(credits)
    if excel_file:
        print(f"\n🎉 SUCCESS! Credits with flat numbers: {excel_file}")
        print("📊 Check the 'Transactions with Flats' sheet for flat-specific transactions")
        print("🏠 Flat numbers are highlighted in yellow in the main sheet")

if __name__ == "__main__":
    main()
