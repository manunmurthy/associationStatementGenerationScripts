"""
Excel Generator Module
Creates formatted Excel files from transaction data
"""

import pandas as pd
from pathlib import Path
from typing import List
from datetime import datetime
import xlsxwriter
from transaction_processor import Transaction


class ExcelGenerator:
    """Class to generate Excel files from transaction data"""
    
    def __init__(self):
        self.output_dir = Path("output")
        self.output_dir.mkdir(exist_ok=True)
    
    def create_excel(self, transactions: List[Transaction], filename_prefix: str) -> Path:
        """
        Create an Excel file from transaction data
        
        Args:
            transactions (List[Transaction]): List of transactions
            filename_prefix (str): Prefix for the output filename
            
        Returns:
            Path: Path to the created Excel file
        """
        # Create filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{filename_prefix}_transactions_{timestamp}.xlsx"
        output_file = self.output_dir / filename
        
        # Convert transactions to DataFrame
        df = self._create_dataframe(transactions)
        
        # Create Excel file with formatting
        self._write_formatted_excel(df, output_file)
        
        return output_file
    
    def _create_dataframe(self, transactions: List[Transaction]) -> pd.DataFrame:
        """Convert transactions to pandas DataFrame"""
        data = []
        running_balance = 0
        
        for transaction in transactions:
            row = transaction.to_dict()
            
            # Calculate running balance if not provided
            if transaction.balance is None:
                if transaction.credit > 0:
                    running_balance += transaction.credit
                elif transaction.debit > 0:
                    running_balance -= transaction.debit
                row['Balance'] = running_balance
            
            data.append(row)
        
        df = pd.DataFrame(data)
        
        # Ensure proper column order
        column_order = ['Date', 'Description', 'Debit', 'Credit', 'Balance', 'Type']
        df = df.reindex(columns=column_order)
        
        return df
    
    def _write_formatted_excel(self, df: pd.DataFrame, output_file: Path):
        """Write DataFrame to Excel with formatting"""
        
        # Create Excel writer object
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Write data to Excel
            df.to_excel(writer, sheet_name='Transactions', index=False)
            
            # Get workbook and worksheet objects for formatting
            workbook = writer.book
            worksheet = writer.sheets['Transactions']
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#4CAF50',  # Green background
                'font_color': 'white',
                'border': 1
            })
            
            currency_format = workbook.add_format({
                'num_format': '$#,##0.00',
                'align': 'right'
            })
            
            date_format = workbook.add_format({
                'num_format': 'dd/mm/yyyy',
                'align': 'center'
            })
            
            debit_format = workbook.add_format({
                'num_format': '$#,##0.00',
                'font_color': '#D32F2F',  # Red for debits
                'align': 'right'
            })
            
            credit_format = workbook.add_format({
                'num_format': '$#,##0.00',
                'font_color': '#388E3C',  # Green for credits
                'align': 'right'
            })
            
            balance_format = workbook.add_format({
                'num_format': '$#,##0.00',
                'bold': True,
                'align': 'right',
                'bg_color': '#F5F5F5'
            })
            
            # Apply header formatting
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Set column widths
            column_widths = {
                'A': 12,  # Date
                'B': 40,  # Description
                'C': 12,  # Debit
                'D': 12,  # Credit
                'E': 12,  # Balance
                'F': 10,  # Type
            }
            
            for col, width in column_widths.items():
                worksheet.set_column(f'{col}:{col}', width)
            
            # Apply conditional formatting for amounts
            # Date column
            worksheet.set_column('A:A', 12, date_format)
            
            # Debit column (C) - Red formatting
            worksheet.set_column('C:C', 12, debit_format)
            
            # Credit column (D) - Green formatting
            worksheet.set_column('D:D', 12, credit_format)
            
            # Balance column (E) - Bold formatting
            worksheet.set_column('E:E', 12, balance_format)
            
            # Add summary section
            self._add_summary_section(worksheet, workbook, df, len(df) + 3)
    
    def _add_summary_section(self, worksheet, workbook, df: pd.DataFrame, start_row: int):
        """Add summary statistics to the worksheet"""
        
        # Summary formatting
        summary_header_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'fg_color': '#2196F3',
            'font_color': 'white'
        })
        
        summary_format = workbook.add_format({
            'num_format': '$#,##0.00',
            'bold': True
        })
        
        # Calculate summary statistics
        total_debits = df['Debit'].replace('', 0).astype(float).sum()
        total_credits = df['Credit'].replace('', 0).astype(float).sum()
        final_balance = df['Balance'].iloc[-1] if len(df) > 0 else 0
        transaction_count = len(df)
        
        # Write summary
        worksheet.write(start_row, 0, "SUMMARY", summary_header_format)
        worksheet.write(start_row + 1, 0, "Total Transactions:")
        worksheet.write(start_row + 1, 1, transaction_count)
        
        worksheet.write(start_row + 2, 0, "Total Credits:")
        worksheet.write(start_row + 2, 1, total_credits, summary_format)
        
        worksheet.write(start_row + 3, 0, "Total Debits:")
        worksheet.write(start_row + 3, 1, total_debits, summary_format)
        
        worksheet.write(start_row + 4, 0, "Final Balance:")
        worksheet.write(start_row + 4, 1, final_balance, summary_format)
    
    def create_simple_excel(self, transactions: List[Transaction], filename_prefix: str) -> Path:
        """
        Create a simple Excel file without advanced formatting
        (Useful as a fallback if xlsxwriter has issues)
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{filename_prefix}_simple_{timestamp}.xlsx"
        output_file = self.output_dir / filename
        
        # Convert to DataFrame
        df = self._create_dataframe(transactions)
        
        # Save simple Excel
        df.to_excel(output_file, sheet_name='Transactions', index=False)
        
        return output_file
