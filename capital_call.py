from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.graphics.shapes import Drawing, Circle
from reportlab.graphics.renderPDF import draw

import random
import pandas as pd 

phone_numbers = {}
account_numbers = {}
abas = {}
swifts = {}


def create_logo(width, height, image_path):
    # Create an Image object with the specified width and height
    logo = Image(image_path, width=width, height=height)

    logo.hAlign = 'LEFT'
    
    return logo


def format_phone_number(num1, num2, num3):
    # Format the string as (xxx)-xxx-xxxx
    formatted_number = "({}) {}-{}".format(num1, num2, num3)
    
    return formatted_number

def create_capital_call_pdf(filename, fund_name_1, fund_name_2, image_path):
    doc = SimpleDocTemplate(filename, pagesize=letter,
    rightMargin=0.5*inch, leftMargin=0.5*inch,
    topMargin=0.5*inch, bottomMargin=0.5*inch)

    story = []

    # Styles
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='CustomHeading1', 
    parent=styles['Heading1'],
    fontSize=20,
    spaceAfter=12))
    styles.add(ParagraphStyle(name='CustomHeading2', 
    parent=styles['Heading2'],
    fontSize=14,
    spaceBefore=6,
    spaceAfter=3))
    styles.add(ParagraphStyle(name='CustomBody', 
    parent=styles['BodyText'],
    fontSize=10,
    textColor=colors.HexColor("#515154"),
    spaceAfter=3))

    # Logo
    logo = create_logo(40, 40, image_path)
    story.append(logo)
    story.append(Spacer(1, 12))

    # Title
    story.append(Paragraph("Capital Call Notice", styles['CustomHeading1']))
    story.append(Spacer(1, 12))

    months = ["September", "October", "November", "December", "January", "February", "March", "April", "May", "June", "July", "August"]
    month_num = random.randrange(12)

    # Header info
    header_info = """
    <b>Date:</b> {month} {day}, {year}<br/>
    <b>To:</b> Limited Partners of {to_name}<br/>
    <b>From: {from_name}, LLC (General Partner)</b> 
    """.format(month=months[month_num], day=random.randrange(1,29), year=2024, to_name=fund_name_1, from_name=fund_name_2)
    story.append(Paragraph(header_info, styles['CustomBody']))
    story.append(Spacer(1, 12))

    # Company Description
    company_description = """
    {fund_name_1} is a private equity investment fund focused on 
    identifying and nurturing high-potential growth companies across various 
    sectors. With a proven track record of strategic investments and value 
    creation, our fund aims to generate superior returns for our limited partners 
    while fostering innovation and sustainable growth in our portfolio companies.
    """.format(fund_name_1=fund_name_1)
    story.append(Paragraph(company_description, styles['CustomBody']))
    story.append(Spacer(1, 12))

    # Capital Call Details
    story.append(Paragraph("Capital Call Details", styles['CustomHeading2']))
    
    total_commitment = random.randrange(800000, 1000000)
    prior_capital_contributions = random.randrange(600000, total_commitment)
    unfunded_commitment = total_commitment - prior_capital_contributions
    current_capital_call = random.randrange(0, unfunded_commitment)
    remaining_unfunded_commitment = unfunded_commitment - current_capital_call

    data = [
        ["Description", "Amount"],
        ["Total Commitment", "${:,}".format(total_commitment)],
        ["Prior Capital Contributions", "${:,}".format(prior_capital_contributions)],
        ["Unfunded Commitment", "${:,}".format(unfunded_commitment)],
        ["Current Capital Call", "${:,}".format(current_capital_call)],
        ["Remaining Unfunded Commitment", "${:,}".format(remaining_unfunded_commitment)],
    ]

    table = Table(data, colWidths=[4*inch, 2.5*inch])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#f5f5f7")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#1d1d1f")),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#f5f5f7")),
        ('TEXTCOLOR', (-1, -2), (-1, -2), colors.HexColor("#0071e3")),
        ('FONTNAME', (-1, -2), (-1, -2), 'Helvetica-Bold'),
        ('LINEBELOW', (0, 0), (-1, -2), 0.5, colors.HexColor("#e5e5e5")),
    ]))
    story.append(table)
    story.append(Spacer(1, 6))

    # Purpose of Capital Call
    story.append(Paragraph("Purpose of Capital Call", styles['CustomHeading2']))
    purpose = """
    • New Portfolio Company Investment: ${:,}<br/>
    • Follow-on Investment: ${:,}<br/>
    • Fund Expenses: ${:,}
    """.format(random.randrange(600000, 1000000), random.randrange(600000, 1000000), random.randrange(600000, 1000000))
    story.append(Paragraph(purpose, styles['CustomBody']))
    story.append(Spacer(1, 6))

    # Payment Instructions
    story.append(Paragraph("Payment Instructions", styles['CustomHeading2']))
    account_number = random.randrange(1000000000, 10000000000)
    ABA = random.randrange(100000000, 1000000000)

    letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    SWIFT = "{}{}{}{}{}{}{}".format(letters[random.randrange(26)], letters[random.randrange(26)], letters[random.randrange(26)], letters[random.randrange(26)], letters[random.randrange(26)], random.randrange(10), random.randrange(10))
    banks = [
    "JPMorgan Chase",
    "Bank of America",
    "Wells Fargo",
    "Citibank",
    "Goldman Sachs",
    "Morgan Stanley",
    "PNC Financial Services",
    "U.S. Bancorp",
    "TD Bank",
    "Capital One",
    "HSBC",
    "Barclays",
    "Deutsche Bank",
    "Credit Suisse",
    "BNP Paribas",
    "UBS",
    "Santander",
    "Royal Bank of Canada",
    "Scotiabank",
    "ING Group"
    ]

    
    payment = """
    <table>
    <tr><td width="100"><b>Bank:</b></td><td>{bank}</td></tr>
    <tr><td><b>Account:</b></td><td>{fund_name_1}</td></tr>
    <tr><td><b>Account Number:</b></td><td>{account_number}</td></tr>
    <tr><td><b>ABA:</b></td><td>{aba}</td></tr>
    <tr><td><b>SWIFT:</b></td><td>{swift}</td></tr>
    </table>
    """.format(bank = random.choice(banks), fund_name_1=fund_name_1, account_number=account_number, aba=ABA, swift=SWIFT)
    story.append(Paragraph(payment, styles['CustomBody']))
    story.append(Spacer(1, 6))

    # Due Date
    month_num += 1
    if (month_num >= 12):
        month_num = 0

    story.append(Paragraph("<b>Due Date:</b> {month} {day}, 2024".format(month=months[month_num], day=random.randrange(1,29)), styles['CustomBody']))
    story.append(Spacer(1, 6))

    # Contact
    legal_names = [
    "Alexander Johnson",
    "Victoria Smith",
    "Benjamin Clarke",
    "Elizabeth Baker",
    "Christopher Miller",
    "Sophia Davis",
    "Matthew Wilson",
    "Isabella Moore",
    "William Taylor",
    "Charlotte Anderson",
    "Nicholas Thomas",
    "Catherine Brown",
    "Jonathan Harris",
    "Margaret Martin",
    "Daniel Thompson",
    "Rebecca Lee",
    "Michael Scott",
    "Alexandra White",
    "Robert Lewis",
    "Katherine Roberts"
    ]


    contact_role = "IR Manager"

    contact = """
    <b>Contact:</b> {contact_name}, {contact_role}<br/>
    {contact_name}@{email}.com | {number}
    """.format(contact_name=random.choice(legal_names), contact_role=contact_role,email=fund_name_2.lower().replace(" ", ""), number=format_phone_number(random.randrange(100,1000), random.randrange(100,1000), random.randrange(1000, 10000)))
    story.append(Paragraph(contact, styles['CustomBody']))

    doc.build(story)

if __name__ == "__main__":

    """
    print("Enter fund_name_1: ")
    fund_name_1 = input()
    print("Enter fund_name_2: ")
    fund_name_2 = input()
    print("Enter logo image path: ")
    image_path = input()
    print("Enter output pdf name: ")
    output_pdf_name = input()
    """

    """
    fund_name_1 = "XYZ Growth Fund III, L.P. 2"
    fund_name_2 = "ABC Capital Management 2"
    image_path = "wide_logo.PNG"
    output_pdf_name = "capital_call_notice"
    
    create_capital_call_pdf(output_pdf_name + ".pdf", fund_name_1, fund_name_2, image_path)
    """


    df = pd.read_excel("data.xlsx")

    index = 0

    investors = df["Investor Names"]

    for investor in investors:
        fund_name = df.at[index, "Fund Names"]
        logo = df.at[index, "Logo"]
        output_pdf_name = investor+fund_name
        create_capital_call_pdf(output_pdf_name + ".pdf", investor, fund_name, logo)
        index += 1