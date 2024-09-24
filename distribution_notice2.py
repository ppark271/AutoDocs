from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, Frame, Spacer, Image, Table, TableStyle
from reportlab.lib import colors

def create_pdf():
    pdf_path = "output.pdf"
    c = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter
    styles = getSampleStyleSheet()

    # Add logo and confidential notice
    logo = "logo.png"  # Replace with the path to the logo image
    c.drawImage(logo, 40, height - 100, width=200, preserveAspectRatio=True, mask='auto')
    c.setFont("Helvetica-Bold", 16)
    c.drawRightString(width - 40, height - 60, "CONFIDENTIAL")
    
    # Addressing details
    address_frame = Frame(40, height - 240, width - 80, 100, showBoundary=0)
    address_content = [
        Paragraph("To: <b>Ben Magleby</b>", styles['Normal']),
        Paragraph("From: <b>Level Structured Capital Associates II, LLC</b>", styles['Normal']),
        Paragraph("RE: Level Structured Capital II (GP), L.P.- Net Distribution #1", styles['Normal']),
        Paragraph("Date: June 7, 2023", styles['Normal']),
    ]
    for para in address_content:
        address_frame.addFromList([para, Spacer(0, 10)], c)
    
    # Main content text
    text = """<b>Net Distribution #1: Distributable Proceeds; net of contributions for Investments, Organizational Expenses, and Partnership Expenses</b><br/><br/>
    Level Structured Capital II (GP), L.P. ("LSCI II GP") is making its first net distribution with respect to its investment in Level Structured Capital II, L.P. ("LSCI II"). This net distribution covers distributions of investment proceeds from LSCI II to LSCI II (GP) from inception to date and is net of investments and expenses for which capital has been called from LSCI II (GP) by LSCI II from inception to date."""
    content_frame = Frame(40, 320, width - 80, height - 300, showBoundary=0)
    content_frame.addFromList([Paragraph(text, styles['Normal'])], c)

    # Financial summary table
    data = [
        ["Allocation of your Distributable Proceeds:", "Current", "Cumulative"],
        ["Committed Capital", "$ -", "$ -"],
        ["Carried Interest", "$ 63", "$ 63"],
        ["Total Distribution", "$ 63", "$ 63"],
        ["Capital Call", "Your Share of GP Gross Contributions", "$ -", "$ -"],
        ["Your Share of GP Deemed Contributions", "$ -", "$ -"],
        ["Your Share of GP Standalone Expenses", "$ -", "$ -"],
        ["Total Capital Call", "$ -", "$ -"],
        ["Net Distribution", "$ 63", "$ 63"],
        ["Less: Equalization Interest", "$ -", "$ -"],
        ["Net Amount Payable", "$ 63", "$ 63"],
        ["Total Estimated Taxable Income", "$ 63", "$ 63"],
        ["Your Remaining Commitment after this Net Distribution", "Capital Commitment", "$ -", "$ -"],
        ["Cumulative Contributions", "$ -", "$ -"],
        ["Remaining Commitment", "$ -", "$ -"],
    ]
    table = Table(data, colWidths=[200, 100, 100])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('SPAN', (0, 0), (-1, 0)),  # Span across all columns for header
        ('SPAN', (0, 1), (0, 2)),  # Span for 'Committed Capital'
        ('SPAN', (0, 3), (0, 4)),  # Span for 'Total Distribution'
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    table.wrapOn(c, width, height)
    table.drawOn(c, 40, 80)

    c.save()

if __name__ == "__main__":
    create_pdf()
