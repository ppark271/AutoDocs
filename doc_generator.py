# Import necessary libraries for PDF generation, data handling, and user interaction
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, legal
from reportlab.lib.pagesizes import A1, A2, A3, A4, A5, A6, landscape
from reportlab.platypus import SimpleDocTemplate
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, Flowable
)
from io import BytesIO
import PyPDF2
from reportlab.pdfgen import canvas
import random  # For generating random data
import pandas as pd  # For reading data from Excel files
from datetime import datetime  # For current date and time
import os  # For file path operations
from tkinter import Tk, filedialog  # For file and directory selection dialogs
import re  # For regular expressions
import shutil
import fitz

from documents.capital_call import *
from documents.quarterly_update import *
from documents.gp_report import *
from documents.wire_instruction import *
from documents.distribution_notice import *

from documents.utils import *



# Main execution block
if __name__ == "__main__":
    # Hide the root window of Tkinter
    root = Tk()
    root.withdraw()

    # Prompt the user to select the Excel file containing the data
    print("Please select the Excel file containing the data:") 
    excel_file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )

    if not excel_file_path:
        print("No file selected. Exiting.")
        exit()

    # Read data from the selected Excel file
    df = pd.read_excel(excel_file_path, sheet_name = "Investor")

    # Prompt the user to select the output directory
    print("Please select the directory where the PDFs will be saved:")
    output_directory = filedialog.askdirectory(title="Select Output Directory")

    if not output_directory:
        print("No directory selected. Exiting.")
        exit()

    print("Select Document type: ")
    print("1. Cap Call")
    print("2. K-Document")
    print("3. Quarterly Update")
    print("4. GP Report")
    print("5. Wire Instruction Confirmation")
    print("6. Distribution Notice")

    document_type = int(input())

    while (document_type < 0 or document_type >= 7):
        print("Invalid document_type, try again")
        print("Select Document type: ")
        print("1. Cap Call")
        print("2. K-Document")
        print("3. Quarterly Update")
        print("4. GP Report")
        print("5. Wire Instruction Confirmation")
        print("6. Distribution Notice")
        document_type = int(input())

    # Get current quarter and year for filename
    now = datetime.now()
    quarter = (now.month - 1) // 3 + 1
    quarter_str = f"Q{quarter} {now.year - 1}"

    logo_path = r"C:\Users\ppark\OneDrive - GP Fund Solutions, LLC\Desktop\doc_generator\aea-logo.PNG"

    funds = {}
    fund_names = set()

    # Iterate over each row in the DataFrame to gather all the funds
    for index, row in df.iterrows():
        fund_names.add(str(row["Fund Name"]))
    
    fund_names = list(fund_names)

    # Iterate over each row in the DataFrame
    for index, row in df.iterrows():
        investing_entity_name = str(row["Fund Name"]) 
        investor_code = str(row["Investor Code"])  
        legal_name = str(row["Legal Name"]) 
        address_1 = str(row["Address 1"])
        address_2 = str(row["Address 2"])
        city = str(row["City"])
        state = str(row["State"])
        zip_code = str(row["Zip"])
        country = str(row["Country"])
        tax_id = str(row["Tax ID"])
        fund_name = str(row["Fund Name"])

        # Check if the logo path is relative or absolute
        if not os.path.isabs(logo_path):
            # If relative, construct the full path relative to the Excel file's directory
            logo_path = os.path.join(os.path.dirname(excel_file_path), logo_path)

        # Verify that the logo file exists
        if not os.path.isfile(logo_path):
            print(f"Logo file '{logo_path}' not found. Skipping {legal_name}.")
            continue

        # Sanitize investor code and legal name for filenames
        investor_code_safe = sanitize_filename(investor_code)
        fund_code_safe = sanitize_filename(investor_code[4:])
        legal_name_safe = sanitize_filename(legal_name)
        output_pdf_path = ""

        text_to_add = {
            "logo" : logo_path,
            "legal_name" : legal_name,
            "footer" : f"{fund_name}, {address_1}, {city}, {state} {zip_code}",
            "date" : datetime.now().strftime("%B %d, %Y"),
            "fund_name" : fund_name,
            "address_1" : address_1
        }

        # Try-except block to catch exceptions
        try:
            #capital call
            if document_type == 1:
                output_pdf_name = f"{investor_code_safe}_{legal_name_safe} - {fund_name} - Capital Call - {quarter_str}.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                create_capital_call_pdf(
                    output_pdf_path,
                    investing_entity_name,
                    legal_name,
                    logo_path
                )
                
            #k1 document
            elif document_type == 2:
                output_pdf_name = f"{investor_code_safe}_{fund_name}_{legal_name_safe}-K1-2023.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                shutil.copy("k1-filled-flat.pdf", output_pdf_path)

            #quarterly update
            elif document_type == 3:
                #if fund has already been encountered, skip it
                if (fund_name in funds):
                    continue
                funds[fund_name] = 1

                output_pdf_name = f"{fund_code_safe} Quarterly Update Page - {quarter_str}.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                create_quarterly_update_pdf(
                    output_pdf_path,
                    investing_entity_name,
                    legal_name,
                    logo_path
                )

                texts_with_positions = [
                    (fund_name, (30, 760))
                ]

                output_pdf_name2 = f"{fund_code_safe} Quarterly Update Page2 - {quarter_str}.pdf"
                output_pdf_path2 = os.path.join(output_directory, output_pdf_name2)
                add_multiple_texts_to_existing_pdf("resized_output.pdf", output_pdf_path2, texts_with_positions)

                output_pdf_name_final = f"{fund_name}_{fund_code_safe}_Quarterly_Update - {quarter_str}.pdf"
                output_pdf_path_final = os.path.join(output_directory, output_pdf_name_final)

                # Step 1: Open the two PDFs
                with open(output_pdf_path, "rb") as file1, open(output_pdf_path2, "rb") as file2:
                    reader1 = PyPDF2.PdfReader(file1)
                    reader2 = PyPDF2.PdfReader(file2)

                    # Step 2: Create a PdfWriter object to hold the combined PDFs
                    writer = PyPDF2.PdfWriter()

                    # Step 3: Add pages from the first PDF
                    for page_num in range(len(reader1.pages)):
                        page = reader1.pages[page_num]
                        writer.add_page(page)

                    # Step 4: Add pages from the second PDF
                    for page_num in range(len(reader2.pages)):
                        page = reader2.pages[page_num]
                        writer.add_page(page)

                    # Step 5: Write the combined PDF to a new file
                    with open(output_pdf_path_final, "wb") as output_file:
                        writer.write(output_file)

                # Step 6: Delete the original PDFs
                os.remove(output_pdf_path)
                os.remove(output_pdf_path2)
                
            #gp report
            elif document_type == 4:
                output_pdf_name = f"{fund_code_safe} GP Report - {quarter_str}.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                footer = f"{investing_entity_name}, {address_1}, {city}, {state}, {zip_code}"
                create_gp_report_pdf(
                    output_pdf_path,
                    investing_entity_name,
                    legal_name,
                    logo_path,
                    fund_names,
                    footer)

            #wire instruction confirmation
            elif document_type == 5:
                output_pdf_name = f"{fund_name}_{investor_code_safe}_{legal_name_safe} Wire Instructions - {quarter_str}.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                create_wire_instruction_pdf(text_to_add, output_pdf_path)

            #distribution notice
            elif document_type == 6:
                output_pdf_name = f"{fund_name}_{investor_code_safe}_{legal_name_safe} Distribution Notice - {quarter_str}.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                create_distribution_notice_pdf(r"C:\Users\ppark\OneDrive - GP Fund Solutions, LLC\Desktop\doc_generator\templates\Distribution Notice #1.pdf", text_to_add, output_pdf_path)

            print(f"Generating PDF for {legal_name} at {output_pdf_path}")
                
        except PermissionError as e:
            print(f"Failed to write PDF for {legal_name}: {e}")
        except Exception as e:
            print(f"An error occurred while generating PDF for {legal_name}: {e}")

    print("PDF generation complete.")
