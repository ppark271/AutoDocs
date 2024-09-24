import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO
import pandas as pd 


def add_multiple_texts_to_existing_pdf(input_pdf, output_pdf, texts_with_positions):
    # Step 1: Create a PDF with all the texts you want to add
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    
    # Loop through the list of texts and positions
    for text, position in texts_with_positions:
        x, y = position
        can.setFont("Helvetica", 8)
        can.drawString(x, y, text)  # Add each text at the specified (x, y) position
    
    can.save()
    
    # Move to the beginning of the BytesIO buffer
    packet.seek(0)

    # Step 2: Read the existing PDF
    existing_pdf = PyPDF2.PdfReader(input_pdf)

    # Step 3: Read the new PDF (with the added text)
    new_pdf = PyPDF2.PdfReader(packet)

    # Step 4: Merge the new text with the existing PDF
    output = PyPDF2.PdfWriter()

    # Iterate through each page of the existing PDF and merge the text
    for i in range(len(existing_pdf.pages)):
        page = existing_pdf.pages[i]
        if i == 0:  # Assuming you want to add text to the first page
            page.merge_page(new_pdf.pages[0])  # Merge the text PDF onto the first page
        output.add_page(page)

    # Step 5: Save the result to a new PDF file
    with open(output_pdf, "wb") as output_file:
        output.write(output_file)

df = pd.read_excel("AEA GP CRM Upload.xlsx", sheet_name = "Investor")
legal_name = str(df.at[1, "Legal Name"])
address = str(df.at[1, "Address 1"])
city = str(df.at[1, "City"])
state = str(df.at[1, "State"])
zip_code = str(df.at[1, "Zip"])

legal_name2 = str(df.at[30, "Legal Name"])
address2 = str(df.at[30, "Address 1"])
city2 = str(df.at[30, "City"])
state2 = str(df.at[30, "State"])
zip_code2 = str(df.at[30, "Zip"])


# Example usage:
partnership_info = legal_name + " " + address + " " + city + " " + state + " " + zip_code
partner_info = legal_name2 + " " + address2 + " " + city2 + " " + state2 + " " + zip_code2
texts_with_positions = [
    (partnership_info, (80, 580)),
    (partner_info, (80, 475))
    # Add more text and positions as needed
]

add_multiple_texts_to_existing_pdf("k-1-document.pdf", "output.pdf", texts_with_positions)
