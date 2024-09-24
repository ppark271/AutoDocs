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

# List of banks to choose from
banks = [
    "JPMorgan Chase", "Bank of America", "Wells Fargo", "Citibank",
    "Goldman Sachs", "Morgan Stanley", "PNC Financial Services",
    "U.S. Bancorp", "TD Bank", "Capital One", "HSBC", "Barclays",
    "Deutsche Bank", "Credit Suisse", "BNP Paribas", "UBS",
    "Santander", "Royal Bank of Canada", "Scotiabank", "ING Group"
]

def generate_account_number():
    return random.randrange(1_000_000_000, 10_000_000_000)
    
def generate_ABA():
    return random.randrange(100_000_000, 1_000_000_000)

def generate_SWIFT():
    letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    
    return ''.join(random.choice(letters) for _ in range(5)) + \
        ''.join(str(random.randrange(10)) for _ in range(2))

def create_logo(max_width, max_height, image_path):
    """
    Creates an Image object for the logo that fits within the specified
    dimensions while maintaining its aspect ratio.
    """
    # Create an Image object using the provided image path
    logo = Image(image_path)
    logo.hAlign = 'LEFT'  # Align the logo to the left

    # Get the original dimensions of the image
    image_width, image_height = logo.wrap(0, 0)

    # Calculate the scaling factor to fit the image within the max dimensions
    scale = min(max_width / image_width, max_height / image_height, 1.0)

    # Apply the scaling factor
    logo.drawWidth = image_width * scale
    logo.drawHeight = image_height * scale

    return logo


# Function to format a phone number in the (xxx) xxx-xxxx format
def format_phone_number(num1, num2, num3):
    """
    Formats a phone number into the (xxx) xxx-xxxx format.
    """
    formatted_number = "({}) {}-{}".format(num1, num2, num3)
    return formatted_number

# Function to create a horizontal divider line
class DividerLine(Flowable):
    """
    A custom Flowable to draw a horizontal divider line.
    """
    def __init__(self, width, thickness=0.5, color=colors.lightgrey):
        super().__init__()
        self.width = width
        self.thickness = thickness
        self.color = color

    def draw(self):
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 0, self.width, 0)

# Function to sanitize filenames by removing invalid characters
def sanitize_filename(filename):
    """
    Removes invalid characters from filenames using regex.
    """
    # Define a regex pattern for invalid characters
    invalid_chars = r'[<>:"/\\|?*]'
    # Replace invalid characters with an underscore
    return re.sub(invalid_chars, '_', filename)

def create_gp_report_pdf2(filename, investing_entity_name, legal_name, image_path):
    """
    Generates a gp report PDF document with the given filename and content.

    Parameters:
    - filename: Name of the output PDF file.
    - investing_entity_name: Name of the investing entity (fund).
    - legal_name: Legal name of the investor.
    - image_path: Path to the logo image file.
    """
    # Create a document template with specified margins and page size
    doc = SimpleDocTemplate(
        filename, pagesize=landscape(A3),
        rightMargin=0.75 * inch, leftMargin=0.75 * inch,
        topMargin=0.75 * inch, bottomMargin=0.75 * inch
    )

    story = []  # List to hold the flowable elements for the PDF

    # Define styles for the document
    styles = getSampleStyleSheet()

    # Update default style with built-in font
    styles['Normal'].fontName = 'Helvetica'
    styles['Normal'].fontSize = 10
    styles['Normal'].leading = 12

    # Custom styles for headings and body text
    styles.add(ParagraphStyle(
        name='CustomHeading1',
        fontName='Helvetica-Bold',
        fontSize=16,
        spaceAfter=10,
        textColor=colors.HexColor("#333333"),
    ))
    styles.add(ParagraphStyle(
        name='CustomHeading2',
        fontName='Helvetica-Bold',
        fontSize=12,
        spaceBefore=14,
        spaceAfter=4,
        textColor=colors.HexColor("#333333"),
    ))
    styles.add(ParagraphStyle(
        name='CustomBody',
        fontName='Helvetica',
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#515154"),
        spaceAfter=10,
    ))
    styles.add(ParagraphStyle(
        name='CustomEmphasis',
        parent=styles['CustomBody'],
        fontName='Helvetica-Bold',
        textColor=colors.HexColor("#000000"),
    ))
    # Style for table cells
    table_cell_style = ParagraphStyle(
        name='TableCell',
        fontName='Helvetica',
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#515154"),
    )
    table_cell_bold_style = ParagraphStyle(
        name='TableCellBold',
        parent=table_cell_style,
        fontName='Helvetica-Bold',
    )

    # Add Logo to the document with aspect ratio maintained
    logo = create_logo(80, 40, image_path)  # Adjust max dimensions as needed
    story.append(logo)
    story.append(Spacer(1, 10))  # Add space after the logo

    # Use the current date for the notice
    current_date = datetime.now().strftime("%B %d, %Y")  # Format: Month Day, Year

    # Create header information with date, recipient, and sender details
    header_info = f"""
    <b>General Partners of the Fund(s)</b><br/>
    Unaudited Capital Account Statement<br/>
    As of and life to date {current_date}<br/>
    <br/>
    <b>{legal_name}</b>
    """
    story.append(Paragraph(header_info.strip(), styles['CustomBody']))

    # Prepare data for the financial table with emphasis on key numbers
    data = [
    ]

    #29
    row_titles = ["Commitment Summary", "Investor commitment", "Gross contributions", "Remaining commitment", "", 
    "Waiver Analysis", "Deemed contributions", "Realized special profit", "Remaining special profit", "", 
    "Life to Date Book Capital Account", "Capital contributions", "Net investment distributions", "Net income / (loss) & syndication costs",
    "Net realized investment gain / (loss)", "Net unrealized investment gain / (loss)", "Net carried interest distributions", "Net realized carried interest",
    "Net unrealized carried interest", "Ending Capital Account Balance", "",
    "Summary", "Investment distributions", "Carried interest distributions", "Total distributions", "Investment Current Value", "Carried interest Current Value", 
    "Total Current Value", "Total Value (Capital + Distributions)"]

    currency_col = ["", "$", "", "$", "", 
    "", "$", "", "$", "", 
    "", "$", "", "",
    "", "", "", "",
    "", "$", "",
    "", "$", "", "$", "$", "", 
    "$", "$"]

    fund1_num1 = random.randrange(1000, 9999)

    fund1 = ["FUND1", "-", "-", "-", "", 
    "", "-", "-", "-", "", 
    "", "-", "-", "-",
    "-", "-", "-", "-",
    f"{fund1_num1}", f"{fund1_num1}", "",
    "", "-", "-", "-", "-", f"{fund1_num1}", 
    f"{fund1_num1}", f"{fund1_num1}"]

    fund2_num1 = random.randrange(10, 999)
    fund2_num2 = random.randrange(1000, 9999)

    fund2 = ["FUND2", "-", "-", "-", "", 
    "", "-", "-", "-", "", 
    "", "-", "-", "-",
    "-", "-", f"({fund2_num1})", f"{fund2_num1}",
    f"{fund2_num2}", f"{fund2_num2}", "",
    "", "-", f"{fund2_num1}", f"{fund2_num1}", "-", f"{fund2_num2}", 
    f"{fund2_num2}", f"{fund2_num2}"]

    fund2[28] = str(int(fund2[27]) + int(fund2[24]))

    fund_info = []
    fund_info.append(
        fund1
    )
    fund_info.append(
        fund2
    )

    total = ["TOTAL", "-", "-", "-", "", 
    "", "-", "-", "-", "", 
    "", "-", "-", "-",
    "-", "-", "(63)", "63",
    "3419", "3419", "",
    "", "-", "63", "63", "-", "3419", 
    "3419", "3482"]

    for i in range(len(fund1)):
        if (fund1[i].isdigit() and fund2[i].isdigit()):
            total[i] = str(int(fund1[i]) + int(fund2[i]))
        elif (fund1[i].isdigit()):
            total[i] = fund1[i]
        elif (fund2[i].isdigit()):
            total[i] = fund2[i]
        elif (fund1[i][1:-1].isdigit()):
            total[i] = fund1[i]
        elif (fund2[i][1:-1].isdigit()):
            total[i] = fund2[i]


    for index in range(len(row_titles)):
        row = []

        row.append(Paragraph(row_titles[index]))
        row.append(Paragraph(currency_col[index]))

        for j in range(len(fund_info)):
            if index >= len(fund_info[j]):
                row.append(Paragraph(""))
                continue
            row.append(Paragraph(fund_info[j][index]))
        
        row.append(Paragraph(total[index]))

        data.append(row)

    # Create and style the financial table
    col_widths = [4*inch, .5 * inch]
    for i in range(len(fund_info)):
        col_widths.append(inch)

    table = Table(data, colWidths=col_widths, hAlign='LEFT', repeatRows=1)

    style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),  # Using Helvetica for all cells
        ('FONTSIZE', (0,0), (-1,-1), 8),  # Set font size to 8 for all cells
        ('BOX', (0,0), (-1,-1), 2, colors.black),
        ('GRID', (0,1), (-1,-1), 1, colors.black)
    ])
    table.setStyle(style)

    story.append(table)

    story.append(Spacer(1, 10))

    address = "Level Equity Management, LLC, Two Grand Central Tower, 140 East 45th Street, 42nd Floor, New York, NY 10017"
    story.append(Paragraph(address, styles['CustomHeading2']))
    
    # Build the PDF document
    doc.build(story)


if __name__ == "__main__":
    output_pdf_path = r"C:\Users\ppark\OneDrive - GP Fund Solutions, LLC\Desktop\doc_generator\GP REPORT TESTING.pdf"

    create_gp_report_pdf2(
        output_pdf_path,
        "AEA",
        "Joe Bruine",
        r"C:\Users\ppark\OneDrive - GP Fund Solutions, LLC\Desktop\doc_generator\aea-logo.png"
    )