"""
PDF Parser Module
Handles extraction of text content from PDF bank statements
"""

import pdfplumber
from pathlib import Path


class PDFParser:
    """Class to handle PDF text extraction"""
    
    def __init__(self):
        """Initialize the PDF parser"""
        self.supported_formats = ['.pdf']
    
    def extract_text(self, file_path: Path) -> str:
        """
        Extract text from a PDF file
        
        Args:
            file_path (Path): Path to the PDF file
            
        Returns:
            str: Extracted text content from the PDF
            
        Raises:
            FileNotFoundError: If the PDF file doesn't exist
            Exception: If PDF processing fails
        """
        if not file_path.exists():
            raise FileNotFoundError(f"PDF file not found: {file_path}")
        
        if file_path.suffix.lower() not in self.supported_formats:
            raise ValueError(f"Unsupported file format: {file_path.suffix}")
        
        try:
            text_content = ""
            
            # Open and extract text from PDF
            with pdfplumber.open(file_path) as pdf:
                print(f"📄 PDF has {len(pdf.pages)} pages")
                
                for page_num, page in enumerate(pdf.pages, 1):
                    print(f"   Processing page {page_num}...")
                    page_text = page.extract_text()
                    
                    if page_text:
                        text_content += page_text + "\n"
                    else:
                        print(f"   ⚠️ No text found on page {page_num}")
            
            if not text_content.strip():
                raise Exception("No text could be extracted from the PDF")
            
            return text_content
            
        except Exception as e:
            raise Exception(f"Error processing PDF: {str(e)}")
    
    def preview_text(self, file_path: Path, lines: int = 20) -> str:
        """
        Extract and preview first few lines of PDF text
        
        Args:
            file_path (Path): Path to the PDF file
            lines (int): Number of lines to preview
            
        Returns:
            str: First 'lines' lines of extracted text
        """
        full_text = self.extract_text(file_path)
        text_lines = full_text.split('\n')
        return '\n'.join(text_lines[:lines])
