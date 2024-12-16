import pdfplumber
import openpyxl
import os
from PIL import Image
import pytesseract

# Define paths
pdf_path = 
excel_file = 

# Check if PDF exists
if not os.path.exists(pdf_path):
    print(f"PDF file not found: {pdf_path}")
else:
    # Open the PDF with pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        # Create a new Excel workbook and sheet
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        # Loop through each page and extract table data
        for page_num, page in enumerate(pdf.pages, start=1):
            print(f"Processing page {page_num}...")

            # Check if the page has tables
            tables = page.extract_tables()

            if tables:
                # Write tables to Excel (we'll write each table on a new sheet)
                for table_index, table in enumerate(tables):
                    sheet_name = f"Page {page_num}_Table {table_index + 1}"
                    new_sheet = workbook.create_sheet(sheet_name)

                    # Write rows of the table to t
