import os
from docx2pdf import convert

def generate_pdf(docx_file="vysledny_posudek.docx", pdf_file="vysledny_posudek.pdf"):
    """
    Převede docx_file -> pdf_file pomocí docx2pdf.
    Docx2pdf vyžaduje MS Word (ve Windows).
    """
    if not os.path.exists(docx_file):
        print(f"Chyba: Soubor {docx_file} neexistuje, nelze převést na PDF.")
        return

    convert(docx_file, pdf_file)
    print(f"Soubor {pdf_file} byl vytvořen (PDF).")
