from documents.utils import *

def create_k1_document_pdf(fund_name, output_pdf_path):
    input_pdf = r"C:\Users\ppark\OneDrive - GP Fund Solutions, LLC\Desktop\doc_generator\documents\k1_template.pdf"
    texts_with_positions = [
        (fund_name + " 1456 Sierra Ridge Drive Fresno CA 93711", (40, 580)),
    ]


    add_multiple_texts_to_existing_pdf(input_pdf, output_pdf_path, texts_with_positions)

