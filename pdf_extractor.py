import PyPDF2

def extract_form_fields(pdf_path):
    """
    Otevře PDF formulář (AcroForm) a vyčte formulářová pole.
    Vrátí slovník {název_pole: hodnota}.
    """
    with open(pdf_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        fields = reader.get_form_text_fields() or {}
        return fields

def parse_personal_data(fields):
    """
    Z 'fields' vytáhne konkrétní údaje a vrátí je ve slovníku.
    Uprav názvy klíčů podle skutečného PDF.
    """
    jmeno_pdf = fields.get("jmeno", "")
    narozeni_pdf = fields.get("narozeni", "")
    # Odstraní lomítko a vše za ním, pokud je přítomno (např. "30.1.1995 / Zlín")
    if "/" in narozeni_pdf:
        narozeni_pdf = narozeni_pdf.split("/")[0].strip()

    adresa_pdf = fields.get("trvale-bydliste-ulice", "")
    mesto_pdf = fields.get("trvale-bydliste-mesto", "")
    psc_pdf = fields.get("trvale-bydliste-psc", "")

    full_adresa = f"{adresa_pdf}, {mesto_pdf}".strip(", ")

    print(">>> parse personal data výsledky:")
    print("Jméno:", jmeno_pdf)
    print("Narození:", narozeni_pdf)
    print("Adresa:", full_adresa)
    print("PSČ:", psc_pdf)

    return {
        "jmeno": jmeno_pdf,
        "narozeni": narozeni_pdf,
        "adresa": full_adresa,
        "psc": psc_pdf
    }
