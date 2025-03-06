from docxtpl import DocxTemplate
import datetime
import os
import sys

def resource_path(relative_path):
    """
    Vrátí absolutní cestu k resource souboru.
    Pokud je aplikace zabalená PyInstallerem, použije se sys._MEIPASS.
    Jinak se vezme aktuální adresář.
    """
    try:
        base_path = sys._MEIPASS  # Když je zabalené PyInstallerem
    except Exception:
        base_path = os.path.abspath(".")  # Normální běh
    return os.path.join(base_path, relative_path)

def generate_docx(
    jmeno, narozeni, adresa, psc,
    pozice, pobocka, konzultant,
    pozice_popis, popis_pozice,
    faktor_values,
    output_file="vysledny_posudek.docx"
):
    """
    Vezme data z textových polí a doplní je do LP_template.docx pomocí placeholderů.
    Šablona LP_template.docx je ve stejné složce jako EXE (po zabalení).
    """
    template_path = resource_path("LP_template.docx")  # <- Bez '_internal'

    doc = DocxTemplate(template_path)
    datum_dnes = datetime.date.today().strftime("%d.%m.%Y")

    context = {
        "jmeno": jmeno,
        "narozeni": narozeni,
        "adresa": adresa,
        "psc": psc,
        "pozice": pozice,
        "pobocka": pobocka,
        "konzultant": konzultant,
        "datum": datum_dnes,
        "pozice_popis": pozice_popis,
        "popis_pozice": popis_pozice,

        "faktor1": faktor_values.get("faktor1", ""),
        "kategorie1": faktor_values.get("kategorie1", ""),
        "faktor2": faktor_values.get("faktor2", ""),
        "kategorie2": faktor_values.get("kategorie2", ""),
        "faktor3": faktor_values.get("faktor3", ""),
        "kategorie3": faktor_values.get("kategorie3", ""),
        "faktor4": faktor_values.get("faktor4", ""),
        "kategorie4": faktor_values.get("kategorie4", ""),
        "faktor5": faktor_values.get("faktor5", ""),
        "kategorie5": faktor_values.get("kategorie5", ""),
        "faktor6": faktor_values.get("faktor6", ""),
        "kategorie6": faktor_values.get("kategorie6", ""),
    }

    doc.render(context)
    doc.save(output_file)
    print(f"Soubor {output_file} byl vytvořen (DOCX).")
