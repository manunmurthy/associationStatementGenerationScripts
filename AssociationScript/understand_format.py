#!/usr/bin/env python3
"""
Step 1: Understand the Bank Statement Format
Let's analyze the HDFC bank statement structure in detail
"""

import sys
import os
from pathlib import Path
import re

def analyze_bank_statement():
    """Carefully analyze the bank statement structure"""
    print("🔍 STEP 1: Understanding Your HDFC Bank Statement")
    print("=" * 55)
    
    try:
        import pdfplumber
        
        # Find PDF file
        input_dir = Path("input")
        pdf_files = list(input_dir.glob("*.pdf"))
        
        if not pdf_files:
            print("❌ No PDF files found in input folder")
            return
        
        pdf_file = pdf_files[0]
        print(f"📂 Analyzing: {pdf_file.name}")
        
        # Extract text from first few pages to understand structure
        with pdfplumber.open(pdf_file) as pdf:
            print(f"📄 PDF has {len(pdf.pages)} pages")
            
            # Analyze first 3 pages in detail
            for page_num in range(min(3, len(pdf.pages))):
                print(f"\n📄 PAGE {page_num + 1} ANALYSIS:")
                print("-" * 40)
                
                page = pdf.pages[page_num]
                page_text = page.extract_text()
                
                if page_text:
                    lines = page_text.split('\n')
                    
                    print(f"Lines on this page: {len(lines)}")
                    
                    # Show all lines with line numbers
                    for i, line in enumerate(lines):
                        line = line.strip()
                        if line:  # Only show non-empty lines
                            print(f"{i+1:3d}: {line}")
                    
                    # Look for transaction patterns on this page
                    print(f"\n🔍 TRANSACTION PATTERNS ON PAGE {page_num + 1}:")
                    transaction_lines = []
                    
                    for i, line in enumerate(lines):
                        # Look for lines that start with date pattern DD/MM/YY
                        if re.match(r'^\d{2}/\d{2}/\d{2}\s', line.strip()):
                            transaction_lines.append((i, line.strip()))
                    
                    if transaction_lines:
                        print(f"Found {len(transaction_lines)} potential transaction lines:")
                        for line_num, line in transaction_lines:
                            print(f"   Line {line_num}: {line}")
                    else:
                        print("No transaction lines found on this page")
                
    except Exception as e:
        print(f"❌ Error: {e}")

def understand_transaction_structure():
    """Analyze a few transaction lines in detail"""
    print("\n\n🔬 STEP 2: Understanding Transaction Structure")
    print("=" * 50)
    
    # Sample transaction lines from your statement
    sample_transactions = [
        "01/12/25 UPI-YASH 0000568231515324 01/12/25 351.00 2,669,063.74",
        "01/12/25 IMPS-533506836089-SAIRAMNANDULA-UTIB-XXX 0000533506836089 01/12/25 3,154.00 2,672,217.74",
        "01/12/25 NEFTCR-ICIC0SF0002-AKSHAR IN12533511041759 01/12/25 3,520.00 2,675,737.74"
    ]
    
    print("Let's break down transaction structure:")
    print("\nSample transactions from your HDFC statement:")
    
    for i, transaction in enumerate(sample_transactions, 1):
        print(f"\n🔍 TRANSACTION {i}:")
        print(f"Raw line: {transaction}")
        
        # Break it down
        parts = transaction.split()
        print("Parts breakdown:")
        for j, part in enumerate(parts):
            print(f"  {j+1:2d}: {part}")
        
        # Identify components
        print("\nIdentified components:")
        
        # Date (first part)
        date_match = re.match(r'^(\d{2}/\d{2}/\d{2})', transaction)
        if date_match:
            print(f"  📅 Date: {date_match.group(1)}")
        
        # Find amounts (numbers with decimals)
        amounts = re.findall(r'[\d,]+\.\d{2}', transaction)
        print(f"  💰 Amounts found: {amounts}")
        
        if len(amounts) >= 2:
            print(f"  💸 Transaction Amount: ₹{amounts[-2]}")
            print(f"  💳 Balance: ₹{amounts[-1]}")
        
        # Description (everything between date and amounts)
        description_part = transaction
        # Remove date
        description_part = re.sub(r'^\d{2}/\d{2}/\d{2}\s+', '', description_part)
        # Remove amounts from end
        for amt in reversed(amounts):
            description_part = description_part.replace(amt, '', 1)
        # Clean up
        description_part = re.sub(r'\s+', ' ', description_part).strip()
        print(f"  📝 Description: {description_part}")
        
        # Determine transaction type
        if 'UPI-' in transaction or 'IMPS-' in transaction:
            if 'CR' in transaction.upper() or 'NEFTCR' in transaction:
                print(f"  🔍 Type: Credit (Money IN)")
            else:
                print(f"  🔍 Type: Debit (Money OUT)")
        
        print("-" * 50)

def create_hdfc_bank_format():
    """Create proper HDFC bank format configuration"""
    print("\n\n⚙️ STEP 3: Creating HDFC Bank Format Configuration")
    print("=" * 55)
    
    hdfc_config = {
        "banks": {
            "hdfc": {
                "name": "HDFC Bank",
                "currency": "INR",
                "date_format": "DD/MM/YY",
                "transaction_patterns": {
                    "date_pattern": r"^\d{2}/\d{2}/\d{2}",
                    "amount_pattern": r"[\d,]+\.\d{2}",
                    "line_structure": "Date Description Reference ValueDate Amount Balance"
                },
                "transaction_types": {
                    "debit_indicators": ["UPI-", "IMPS-", "ATM", "PURCHASE", "WITHDRAWAL", "FEE", "CHARGE"],
                    "credit_indicators": ["NEFTCR", "DEPOSIT", "SALARY", "INTEREST", "CREDIT", "RTGSCR"]
                },
                "amount_extraction": {
                    "transaction_amount_position": -2,
                    "balance_position": -1,
                    "note": "Last amount is balance, second-last is transaction amount"
                }
            }
        }
    }
    
    print("🏦 HDFC Bank Format Analysis:")
    print(f"  Currency: {hdfc_config['banks']['hdfc']['currency']}")
    print(f"  Date Format: {hdfc_config['banks']['hdfc']['date_format']}")
    print(f"  Transaction Pattern: {hdfc_config['banks']['hdfc']['transaction_patterns']['date_pattern']}")
    print("  Debit Indicators:", ", ".join(hdfc_config['banks']['hdfc']['transaction_types']['debit_indicators']))
    print("  Credit Indicators:", ", ".join(hdfc_config['banks']['hdfc']['transaction_types']['credit_indicators']))
    
    # Save to config file
    import json
    config_file = Path("config/hdfc_bank_format.json")
    config_file.parent.mkdir(exist_ok=True)
    
    with open(config_file, 'w') as f:
        json.dump(hdfc_config, f, indent=2)
    
    print(f"\n✅ HDFC configuration saved to: {config_file}")
    
    return hdfc_config

def main():
    """Main analysis function"""
    print("🏦 HDFC BANK STATEMENT ANALYSIS")
    print("=" * 40)
    print("🎯 Goal: Understand the format before building converter")
    print()
    
    # Step 1: Analyze the actual PDF
    analyze_bank_statement()
    
    # Step 2: Understand transaction structure
    understand_transaction_structure()
    
    # Step 3: Create proper configuration
    config = create_hdfc_bank_format()
    
    print("\n" + "="*60)
    print("📋 SUMMARY - What We Learned:")
    print("="*60)
    print("✅ HDFC Bank uses DD/MM/YY date format")
    print("✅ Currency is Indian Rupees (INR)")
    print("✅ Each transaction line has: Date, Description, Reference, ValueDate, Amount, Balance")
    print("✅ Last number is account balance")
    print("✅ Second-last number is transaction amount")
    print("✅ UPI/IMPS without 'CR' = Debit (money out)")
    print("✅ NEFTCR/RTGSCR = Credit (money in)")
    print()
    print("🎯 Next Step: Build a converter specifically for this HDFC format")

if __name__ == "__main__":
    main()
