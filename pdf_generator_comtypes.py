import os
import comtypes.client

def generate_pdf_with_comtypes(
    docx_file="C:\\cesta\\k\\souboru\\vysledny_posudek.docx",
    pdf_file="C:\\cesta\\k\\souboru\\vysledny_posudek.pdf"
):
    """
    Převádí DOCX soubor na PDF pomocí COM rozhraní MS Wordu.
    Ujisti se, že zadáváš absolutní cesty, protože Python není v PATH.
    """
    if not os.path.exists(docx_file):
        print(f"Chyba: Soubor {docx_file} neexistuje, nelze převést na PDF.")
        return

    # Vytvoří COM objekt pro MS Word
    word = comtypes.client.CreateObject('Word.Application')
    # Pokud chceš vidět okno Wordu, odkomentuj níže:
    # word.Visible = True

    # Otevře dokument pomocí absolutní cesty
    doc = word.Documents.Open(os.path.abspath(docx_file))
    
    # Export do PDF (ExportFormat=17 znamená PDF)
    doc.ExportAsFixedFormat(
        OutputFileName=os.path.abspath(pdf_file),
        ExportFormat=17,  # 17 = PDF
        OpenAfterExport=False,
        OptimizeFor=0,    # 0 = kvalita pro tisk
        CreateBookmarks=1 # automatické záložky podle nadpisů
    )
    doc.Close()
    word.Quit()

    print(f"Soubor {pdf_file} byl vytvořen (PDF).")

if __name__ == "__main__":
    generate_pdf_with_comtypes()
