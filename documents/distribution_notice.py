from documents.utils import *

def create_distribution_notice_pdf(filename, text_to_add, output_pdf_path):
    doc = fitz.open(filename)
    page = doc.load_page(0)

    logo_blocker = fitz.Rect(50, 20, 300, 60)
    top_blocker = fitz.Rect(140, 85, 500, 160)
    paragraph_blocker = fitz.Rect(50, 205, 600, 270)

    page.draw_rect(logo_blocker, color=(1, 1, 1), fill=(1, 1, 1))
    page.draw_rect(top_blocker, color=(1, 1, 1), fill=(1, 1, 1))
    page.draw_rect(paragraph_blocker, color=(1, 1, 1), fill=(1, 1, 1))

    #logo
    #legal name
    page.insert_text((140, 95), text_to_add["legal_name"], fontsize=12, color=(0,0,0))
    #fund name
    page.insert_text((140, 117), text_to_add["fund_name"], fontsize=12, color=(0,0,0))
    #fund name re?
    page.insert_text((140, 138), text_to_add["fund_name"], fontsize=12, color=(0,0,0))
    #date
    page.insert_text((140, 158), text_to_add["date"], fontsize=12, color=(0,0,0))
    #paragraph
    paragraph_text = """
    Level Structured Capital II (GP), L.P. (“LSC II GP”) is making its first net distribution with respect to its 
    investment in Level Structured Capital II, L.P. (“LSC II”). This net distribution covers distributions of 
    investment proceeds from LSC II to LSC II (GP) from inception to date and is net of investments and 
    expenses for which capital has been called from LSC II (GP) by LSC II from inception to date
    """

    text_location = fitz.Point(40, 200)  # x=72, y=100 (1 inch margin from top-left corner)
    page.insert_text(text_location, paragraph_text, fontsize=12, fontname="helv")

    doc.save(output_pdf_path)
    doc.close()