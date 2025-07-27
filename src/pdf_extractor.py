"""
ANSI Standards PDF Data Extractor

This script extracts structured data from ANSI standards PDF documents
and exports the results to Excel format.

Author: Muhammed Kutashi
Date: July 2025
"""

import warnings
import PyPDF2
import pandas as pd
import re
import os
from pathlib import Path

warnings.filterwarnings("ignore")

class ANSIStandardsExtractor:
    def __init__(self, pdf_path, output_path="data/output/"):
        """
        Initialize the ANSI Standards Extractor
        
        Args:
            pdf_path (str): Path to the input PDF file
            output_path (str): Directory to save the output Excel file
        """
        self.pdf_path = pdf_path
        self.output_path = Path(output_path)
        self.output_path.mkdir(parents=True, exist_ok=True)
        
        # Initialize data containers
        self.operating_name = []
        self.legal_name = []
        self.website = []
        self.document_name = []
        self.standard_title = []
        self.publishing_date = []
        
    def extract_pdf_text(self):
        """Extract text from all pages of the PDF"""
        try:
            with open(self.pdf_path, 'rb') as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                num_pages = len(pdf_reader.pages)
                
                text = []
                for page_num in range(num_pages):
                    page_obj = pdf_reader.pages[page_num]
                    page_text = page_obj.extract_text()
                    text.append(page_text)
                
                return text
        except FileNotFoundError:
            raise FileNotFoundError(f"PDF file not found: {self.pdf_path}")
        except Exception as e:
            raise Exception(f"Error reading PDF: {str(e)}")
    
    def parse_text_data(self, text):
        """Parse extracted text to structured data"""
        pattern = r"\S+ \("
        
        for i, page_text in enumerate(text):
            # Handle first page differently (skip header lines)
            if i == 0:
                page = page_text.split('\n')[11:-2]
            else:
                page = page_text.split('\n')[:-2]
            
            # Check for website information
            web_test = [x for x in page if 'w: ' in x]
            if not web_test:
                self.website.append('Nil')
            
            # Find Final Action Date indices
            final_action_indices = []
            for idx, string in enumerate(page):
                if 'Final Action Date' in string:
                    final_action_indices.append(idx)
            
            # Process each section
            start = -1
            for idx in final_action_indices:
                sub_list = page[start+1:idx+1]
                start = idx
                
                # Extract publishing date
                val = sub_list[-1].split('|')
                for k in val:
                    if 'Final Action Date' in k:
                        date = k.split('Final Action Date: ')[1].strip()
                        self.publishing_date.append(date)
                
                # Extract operating and legal names
                self._extract_names(sub_list, pattern)
                
                # Extract website
                self._extract_website(sub_list)
                
                # Extract document name and title
                self._extract_document_info(sub_list)
                
                # Ensure data consistency
                self._ensure_data_consistency()
    
    def _extract_names(self, sub_list, pattern):
        """Extract operating and legal names from text section"""
        for line in sub_list:
            if re.match(pattern, line):
                self.operating_name.append(line.split()[0])
                if 'The data in this document is reported' in line:
                    self.legal_name.append(', '.join(line.split()[1:])[:-106])
                else:
                    self.legal_name.append(', '.join(line.split()[1:]))
    
    def _extract_website(self, sub_list):
        """Extract website information from text section"""
        for line in sub_list:
            if 'w: ' in line:
                val = line.split('|')[1].split('w: ')[1].strip()
                self.website.append(val)
    
    def _extract_document_info(self, sub_list):
        """Extract document name and standard title"""
        for idx, line in enumerate(sub_list):
            if (line.startswith('ANSI') or line.startswith('INCITS')) and ',' in line:
                index = idx
                break
        
        doc = ' '.join(sub_list[index:-1]).split(',')[0]
        self.document_name.append(doc)
        
        title = ' '.join(' '.join(sub_list[index:-1]).split(',')[1:]).lstrip(' ')
        self.standard_title.append(title)
    
    def _ensure_data_consistency(self):
        """Ensure all data lists have consistent length"""
        if len(self.operating_name) < len(self.standard_title):
            self.operating_name.append(self.operating_name[-1])
            self.legal_name.append(self.legal_name[-1])
        
        if len(self.website) < len(self.standard_title):
            self.website.append(self.website[-1])
    
    def create_dataframe(self):
        """Create pandas DataFrame from extracted data"""
        data = {
            'Operating Name': self.operating_name,
            'Legal Name': self.legal_name,
            'Website': self.website,
            'Document Name': self.document_name,
            'Standard Title': self.standard_title,
            'Publishing Date': self.publishing_date
        }
        
        df = pd.DataFrame(data)
        df['Legal Name'] = df['Legal Name'].str.replace('[()]', '', regex=True)
        df['American Standard'] = 'American'
        
        return df
    
    def save_to_excel(self, df, filename="ANSI_Extracted.xlsx"):
        """Save DataFrame to Excel file"""
        output_file = self.output_path / filename
        df.to_excel(output_file, index=False)
        print(f"Data saved to: {output_file}")
        return output_file
    
    def process(self, output_filename="ANSI_Extracted.xlsx"):
        """Main processing pipeline"""
        print("Extracting text from PDF...")
        text = self.extract_pdf_text()
        
        print("Parsing text data...")
        self.parse_text_data(text)
        
        print("Creating DataFrame...")
        df = self.create_dataframe()
        
        print(f"Extracted {len(df)} records")
        print(f"DataFrame shape: {df.shape}")
        
        print("Saving to Excel...")
        output_file = self.save_to_excel(df, output_filename)
        
        return df, output_file

def main():
    """Main function to run the extractor"""
    # Example usage
    pdf_path = "data/input/Ansi_Standards_asof_Jan2323.pdf"
    
    if not os.path.exists(pdf_path):
        print(f"Please place your PDF file at: {pdf_path}")
        return
    
    extractor = ANSIStandardsExtractor(pdf_path)
    df, output_file = extractor.process()
    
    print("\nFirst 5 rows of extracted data:")
    print(df.head())
    
    print(f"\nDataFrame info:")
    print(df.info())

if __name__ == "__main__":
    main()
