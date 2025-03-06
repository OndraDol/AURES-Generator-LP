import sys
import os
import platform
import subprocess
import locale
import tempfile
import struct

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout,
    QGroupBox, QPushButton, QLabel, QComboBox,
    QMessageBox, QFileDialog, QLineEdit, QFrame, QCheckBox
)
from PyQt6.QtCore import Qt, QSettings
from qt_material import apply_stylesheet

# Vlastní moduly
from pdf_extractor import extract_form_fields, parse_personal_data
from docx_generator import generate_docx
from pdf_generator_comtypes import generate_pdf_with_comtypes

# Nastavíme locale (pokud to jde)
try:
    locale.setlocale(locale.LC_ALL, 'cs_CZ.UTF-8')
except locale.Error:
    print("Nepodařilo se nastavit české locale 'cs_CZ.UTF-8'.")

###############################################################################
# NOVÝ SEZNAM FAKTORŮ
###############################################################################
FACTOR_OPTIONS = [
    # První blok
    "Zátěž chladem",
    "Fyzická zátěž",
    "Pracovní poloha",
    "Hluk",
    "Vibrace",
    "Prach",
    "Chemické látky a směsi",
    # Oddělovník
    "------------------",
    # Zbytek abecedně
    "Biologické činitele",
    "Činnosti epidemiologicky závažné",
    "Elektrotechnik dle vyhl č.50 / 1978 Sb.",
    "Neionizující záření",
    "Obsluha tlakových nádob a zařízení vys. napětí",
    "Obsluha transport. a vysokozdvižných vozíků",
    "Obsluha transport. vytáhů a jeřábů a vázaní břemen",
    "Práce v noci",
    "Práce ve výškách a nad volnou hloubkou",
    "Psychická zátěž",
    "Řidič v pracovněprávním vztahu do 7,5 tuny",
    "Řidič v pracovněprávním vztahu nad 7,5 tuny",
    "Vyšetření v základním rozsahu",
    "Zátež teplem",
    "Zraková zátěž"
]

# Zbytek (popisy pozic, pobočky, konzultanti) beze změn
position_descriptions = {
    "Call Centrum": "Činnost spojená s jednáním s klienty/zaměstnanci po telefonu, prací na počítači a administrativními úkony. Provoz: dvousměnný, Pracovní doba: 8 hodinové směny (pondělí - neděle) v průměru 40 hodin/týdně.",
    "Manažer/Zástupce manažera": "Administrativní činnost spojená s prací na počítači, s vedením pracovního týmu a jednání s klienty. Součástí práce je řízení služebního vozidla. Provoz: jednosměnný, Pracovní doba: 12 hodinové směny (krátký/dlouhý týden)",
    "Prodejce": "Činnost spojená s prací na počítači a jednání s klienty. Součástí práce je řízení vozidla. Práce je vykonávaná v administrativních budovách a zároveň ve venkovních prostorech. V zimních měsících není překročena hranice čtyř hodin za pracovní dobu, kdyby zaměstnanci byli vystaveni operativní teplotě +4 st. C a nižší. Práce prodejce je spojená se střídáním pobytu v teple a v chladu. Provoz: jednosměnný, Pracovní doba: 12 hodinové směny (krátký/dlouhý týden)",
    "Nákupčí / Testovací technik": "Administrativní činnost spojená s prací na počítači a jednání s klienty. Součástí práce je řízení vozidla. Občasný pohyb ve venkovních prostorech. Provoz: jednosměnný, Pracovní doba: 12 hodinové směny (krátký/dlouhý týden)",
    "Mobilní nákupčí": "Administrativní činnost spojená s prací na počítači a jednání s klienty. Součástí práce je řízení vozidla (více než 60%). Občasný pohyb ve venkovních prostorech. Provoz: jednosměnný, Pracovní doba: 12 hodinové směny (krátký/dlouhý týden)",
    "Specialista zákaznického servisu": "Administrativní činnost spojená s prací na počítači a jednání s klienty.  Provoz: jednosměnný, Pracovní doba: 12 hodinové směny (krátký/dlouhý týden)",
    "Mobilní Specialista zákaznického servisu": "Administrativní činnost spojená s prací na počítači a jednání s klienty. Součástí práce je řízení vozidla.  Provoz: jednosměnný, Pracovní doba: 12 hodinové směny (krátký/dlouhý týden)",
    "Administrativa": "Administrativní činnost spojená s prací na počítači, zakládáním složek a jednání se zaměstnanci/klienty. Provoz: jednosměnný, Pracovní doba: 8 hodinové směny (pondělí - pátek)",
    "Operační tým/Pracovník podpory prodeje/Vedoucí úseku/Mobilní pracovník podpory prodeje": "Drobné práce v dílně, mytí a čištění automobilů, část pracovní doby je práce venku. Součástí práce je řízení vozidla. Práce s chemickými látkami s H314. H340, H350, H372. Provoz: jednosměnný, Pracovní doba: 12 hodinové směny (krátký/dlouhý týden)",
    "Uklízečka": "Provádí úklidové práce, mytí, čištění, vynášení odpadků. Provoz: jednosměnný, Pracovní doba: 8 hodinové směny (pondělí - pátek)",
    "Stock kontrolor": "Převoz automobilů v objektu společnosti, kontrola vozového parku, občasná práce s počítačem. Součástí práce je řízení vozidla. Provoz: jednosměnný, Pracovní doba: 8 hodinové směny (pondělí - pátek)",
    "Řidič do 7,5t/ Kurýr": "Provádí přepravu zákazníků a vozidel. Součástí práce je řízení vozidla. Provoz: jednosměnný, Pracovní doba: 8 hodinové směny (pondělí - pátek)",
    "Řidič nad 7,5t": "Provádí přepravu vozů mezi pobočkami, nakládka a vykládka vozů. Součástí práce je řízení vozidla nad 7,5 tuny. Provoz: jednosměnný, Pracovní doba: 8 hodinové směny (pondělí - pátek)",
    "Technik se zaměřením na opravy detailů": "Provádí drobné opravy detailů vozu. Součástí práce je řízení vozidla. Provoz: jednosměnný, Pracovní doba: 8 hodinové směny (pondělí - pátek)",
    "Mechanik": "Provádí údržbu, opravy a seřizování silničních motorových vozidel. Práce s chemickými látkami s H314. Součástí práce je řízení vozidla. Provoz: jednosměnný, Pracovní doba: 8 hodinové směny (pondělí - pátek)",
    "Výpomoc při přepravě vozů mezi pobočkami": "Provádí přepravu zákazníků a vozidel. Součástí práce je řízení vozidla. Provoz: jednosměnný, DPP spolupráce.",
    "Výpomoc při prodeji vozů": "Činnost spojená s prací na počítači a jednání s klienty. Součástí práce je řízení vozidla. Práce je vykonávaná v administrativních budovách a zároveň ve venkovních prostorech. V zimních měsících není překročena hranice čtyř hodin za pracovní dobu, kdyby zaměstnanci byli vystaveni operativní teplotě +4 st. C a nižší. Práce je spojená se střídáním pobytu v teple a v chladu. Povaha práce nárazová (dohoda o provedení práce), směny maximálně 8hodinové."
}

branch_list = [
    "Brno", "Chomutov", "České Budějovice", "Čestlice",
    "Hradec Králové", "Jihlava", "Kladno", "Kolín",
    "Liberec", "Mladá Boleslav", "Olomouc", "Opava",
    "Ostrava", "Pardubice", "Plzeň", "Praha", "Sokolov",
    "Tábor", "Teplice", "Ústí nad Labem", "Valašské Meziříčí",
    "Zlín", "Znojmo"
]

consultant_list = [
    "Anna Brůčková", "Anna Kučerová", "Alžběta Čermáková", "Blanka Poliaková",
    "Jan Jarma", "Kamila Hušková", "Karolína Kulvaitová", "Miroslava Válková",
    "Ondřej Dolejš"
]

###############################################################################
# Pomocné funkce
###############################################################################
def open_file(file_path: str):
    if os.path.exists(file_path):
        system_name = platform.system().lower()
        try:
            if system_name.startswith('win'):
                os.startfile(file_path)
            elif system_name.startswith('darwin'):
                subprocess.run(['open', file_path])
            else:
                subprocess.run(['xdg-open', file_path])
        except Exception as e:
            print(f"Soubor nelze automaticky otevřít: {e}")

def get_surname(full_name: str) -> str:
    parts = full_name.strip().split()
    if not parts:
        return "nezadano"
    return parts[-1]

def parseOutlookFileName(file_group):
    data = bytes(file_group)
    if len(data) >= 4:
        count = struct.unpack('<I', data[:4])[0]
        if count > 0 and len(data) >= 4 + 592:
            descriptor = data[4:4+592]
            filename = descriptor[76:76+520].decode('utf-16le', errors='ignore').split('\x00')[0]
            return filename
    return "unknown.pdf"

class DragDropFrame(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("dragDropFrame")

        self.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
        self.setAcceptDrops(True)
        self.setFixedSize(220, 90)

        label = QLabel("Přetáhni PDF", self)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout = QVBoxLayout(self)
        layout.addWidget(label)

    def dragEnterEvent(self, event):
        mime = event.mimeData()
        if mime.hasUrls():
            for url in mime.urls():
                if url.toLocalFile().lower().endswith(".pdf"):
                    event.acceptProposedAction()
                    return
        elif mime.hasFormat("FileGroupDescriptorW"):
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event):
        mime = event.mimeData()
        if mime.hasUrls():
            for url in mime.urls():
                local_path = url.toLocalFile()
                if local_path.lower().endswith(".pdf"):
                    main_win = self.parentWidget().parentWidget()
                    if hasattr(main_win, "load_pdf_file"):
                        main_win.load_pdf_file(local_path)
                    QMessageBox.information(self, "Načteno", f"PDF {local_path} bylo načteno (drag & drop).")
        elif mime.hasFormat("FileGroupDescriptorW"):
            file_group = mime.data("FileGroupDescriptorW")
            filename = parseOutlookFileName(file_group)
            if mime.hasFormat("FileContents"):
                file_contents = mime.data("FileContents")
                temp_path = os.path.join(tempfile.gettempdir(), filename)
                with open(temp_path, 'wb') as f:
                    f.write(file_contents)
                main_win = self.parentWidget().parentWidget()
                if hasattr(main_win, "load_pdf_file"):
                    main_win.load_pdf_file(temp_path)
                QMessageBox.information(self, "Načteno", f"PDF {temp_path} bylo načteno (drag & drop z Outlooku).")
        event.acceptProposedAction()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AURES Generátor LP")

        self.settings = QSettings("MyCompany", "AURESGenerátorLP")

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        self.main_layout = QVBoxLayout(main_widget)
        self.main_layout.setContentsMargins(30, 30, 30, 30)
        self.main_layout.setSpacing(20)

        # 1) Horní řada: Drag & Drop, Nadpis, Tlačítko
        top_layout = QHBoxLayout()
        top_layout.setSpacing(30)

        self.drag_drop_frame = DragDropFrame(self)
        top_layout.addWidget(self.drag_drop_frame, alignment=Qt.AlignmentFlag.AlignLeft)

        title_label = QLabel("AURES Generátor LP")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 40px; font-weight: bold; color: #ffffff; border: none; background: none;")
        top_layout.addWidget(title_label, alignment=Qt.AlignmentFlag.AlignCenter)

        self.load_pdf_btn = QPushButton("Nahrát PDF")
        self.load_pdf_btn.setStyleSheet("""
            QPushButton {
                background-color: #D32F2F;
                color: white;
                border-radius: 8px;
                font-size: 14px;
                padding: 12px 24px;
                border: 1px solid transparent;
            }
            QPushButton:hover {
                background-color: #F44336;
            }
            QPushButton:focus {
                outline: none;
                border: 1px solid #ffffff;
            }
        """)
        self.load_pdf_btn.setFixedHeight(50)
        self.load_pdf_btn.clicked.connect(self.on_load_pdf_clicked)
        top_layout.addWidget(self.load_pdf_btn, alignment=Qt.AlignmentFlag.AlignRight)

        self.main_layout.addLayout(top_layout)

        # 2) GroupBox: "Údaje z PDF"
        pdf_box = QGroupBox("Údaje z PDF")
        pdf_box_layout = QGridLayout(pdf_box)
        pdf_box_layout.setSpacing(15)

        pdf_box_layout.addWidget(QLabel("Jméno:"), 0, 0)
        self.jmeno_line = QLineEdit()
        pdf_box_layout.addWidget(self.jmeno_line, 0, 1)

        pdf_box_layout.addWidget(QLabel("Adresa:"), 0, 2)
        self.adresa_line = QLineEdit()
        pdf_box_layout.addWidget(self.adresa_line, 0, 3)

        pdf_box_layout.addWidget(QLabel("Datum narození:"), 1, 0)
        self.narozeni_line = QLineEdit()
        pdf_box_layout.addWidget(self.narozeni_line, 1, 1)

        pdf_box_layout.addWidget(QLabel("PSČ:"), 1, 2)
        self.psc_line = QLineEdit()
        pdf_box_layout.addWidget(self.psc_line, 1, 3)

        self.main_layout.addWidget(pdf_box)

        # 3) GroupBox: Faktory a Kategorie (1–6)
        factor_box = QGroupBox("Faktory a Kategorie (1–6)")
        factor_layout = QGridLayout(factor_box)
        factor_layout.setHorizontalSpacing(30)
        factor_layout.setVerticalSpacing(15)

        self.faktor_combos = []
        self.kategorie_combos = []

        for i in range(6):
            faktor_combo = QComboBox()
            faktor_combo.addItem("")
            faktor_combo.addItems(FACTOR_OPTIONS)

            kategorie_combo = QComboBox()
            kategorie_combo.addItems(["", "Kat. 1", "Kat. 2", "Kat. 3", "Kat. 4"])

            faktor_combo.setMinimumWidth(250)
            kategorie_combo.setMinimumWidth(70)

            self.faktor_combos.append(faktor_combo)
            self.kategorie_combos.append(kategorie_combo)

        self.faktor_combos[0].setCurrentText("Vyšetření v základním rozsahu")
        self.kategorie_combos[0].setCurrentText("Kat. 1")

        factor_layout.addWidget(self.faktor_combos[0], 0, 0)
        factor_layout.addWidget(self.kategorie_combos[0], 0, 1)
        factor_layout.addWidget(self.faktor_combos[3], 0, 2)
        factor_layout.addWidget(self.kategorie_combos[3], 0, 3)

        factor_layout.addWidget(self.faktor_combos[1], 1, 0)
        factor_layout.addWidget(self.kategorie_combos[1], 1, 1)
        factor_layout.addWidget(self.faktor_combos[4], 1, 2)
        factor_layout.addWidget(self.kategorie_combos[4], 1, 3)

        factor_layout.addWidget(self.faktor_combos[2], 2, 0)
        factor_layout.addWidget(self.kategorie_combos[2], 2, 1)
        factor_layout.addWidget(self.faktor_combos[5], 2, 2)
        factor_layout.addWidget(self.kategorie_combos[5], 2, 3)

        self.main_layout.addWidget(factor_box)

        # 4) Pobočka, Pozice, Konzultant
        bottom_layout = QHBoxLayout()
        bottom_layout.setSpacing(20)

        bottom_layout.addWidget(QLabel("Pobočka:"))
        self.pobocka_combo = QComboBox()
        self.pobocka_combo.addItems(sorted(branch_list, key=locale.strxfrm))
        bottom_layout.addWidget(self.pobocka_combo)

        bottom_layout.addWidget(QLabel("Pozice:"))
        self.pozice_line = QLineEdit()
        bottom_layout.addWidget(self.pozice_line)

        bottom_layout.addWidget(QLabel("Konzultant:"))
        self.konzultant_combo = QComboBox()
        self.konzultant_combo.addItems(sorted(consultant_list, key=locale.strxfrm))
        bottom_layout.addWidget(self.konzultant_combo)

        self.remember_checkbox = QCheckBox("Zapamatovat")
        bottom_layout.addWidget(self.remember_checkbox)

        saved_consultant = self.settings.value("consultant", "")
        if saved_consultant:
            index = self.konzultant_combo.findText(saved_consultant)
            if index != -1:
                self.konzultant_combo.setCurrentIndex(index)
                self.remember_checkbox.setChecked(True)

        self.main_layout.addLayout(bottom_layout)

        # 5) Pozice popis
        popis_layout = QHBoxLayout()
        popis_layout.setSpacing(10)

        popis_label = QLabel("Pozice popis:")
        popis_layout.addWidget(popis_label, alignment=Qt.AlignmentFlag.AlignLeft)

        self.pozice_popis_combo = QComboBox()
        sorted_positions = sorted(position_descriptions.keys(), key=locale.strxfrm)
        self.pozice_popis_combo.addItems(sorted_positions)
        popis_layout.addWidget(self.pozice_popis_combo, alignment=Qt.AlignmentFlag.AlignLeft)

        popis_layout.addStretch(1)
        self.main_layout.addLayout(popis_layout)

        # 6) Tlačítka dole
        btn_layout = QHBoxLayout()

        self.gen_docx_btn = QPushButton("Generovat DOCX")
        self.gen_docx_btn.setFixedHeight(50)
        self.gen_docx_btn.setStyleSheet("""
            QPushButton {
                background-color: #007ACC; 
                color: white; 
                font-size: 14px; 
                padding: 12px; 
                border-radius: 8px;
                border: 1px solid transparent;
            }
            QPushButton:hover {
                background-color: #1A73E8;
            }
            QPushButton:focus {
                outline: none;
                border: 1px solid #ffffff;
            }
        """)
        self.gen_docx_btn.clicked.connect(self.on_generate_docx)
        btn_layout.addWidget(self.gen_docx_btn)

        self.gen_pdf_btn = QPushButton("Generovat PDF")
        self.gen_pdf_btn.setFixedHeight(50)
        self.gen_pdf_btn.setStyleSheet("""
            QPushButton {
                background-color: #D32F2F; 
                color: white; 
                font-size: 14px; 
                padding: 12px; 
                border-radius: 8px;
                border: 1px solid transparent;
            }
            QPushButton:hover {
                background-color: #F44336;
            }
            QPushButton:focus {
                outline: none;
                border: 1px solid #ffffff;
            }
        """)
        self.gen_pdf_btn.clicked.connect(self.on_generate_pdf)
        btn_layout.addWidget(self.gen_pdf_btn)

        self.main_layout.addLayout(btn_layout)

        self.resize(900, 600)
        self.set_dark_style()

    def set_dark_style(self):
        """Použijeme styl (qt-material) dark_teal.xml + extra nastavení."""
        apply_stylesheet(QApplication.instance(), theme='dark_teal.xml', extra={'density_scale': '-1'})

        self.setStyleSheet("""
            QWidget {
                background-color: #1E1E1E;
            }
            QGroupBox {
                border: none;
                background-color: transparent;
            }
            QGroupBox::title {
                color: #ffffff;
                font-size: 14px;
            }
            QLabel {
                color: #ffffff;
                font-size: 14px;
                background: none;
            }
            QLineEdit {
                color: #ffffff;
                background-color: #3C3F41;
                border: 1px solid #555;
                padding: 2px;
                border-radius: 5px;
                min-height: 28px;
            }
            QComboBox {
                color: #ffffff;
                background-color: #3C3F41;
                border: 1px solid #555;
                padding: 2px;
                border-radius: 5px;
                min-height: 28px;
            }
            QComboBox QAbstractItemView {
                background-color: #3C3F41; 
                color: #ffffff;  
                border: 1px solid #555;
            }
            QComboBox QAbstractItemView::item {
                background-color: #3C3F41;
                color: #ffffff;
            }
            QComboBox QAbstractItemView::item:hover {
                background-color: #666666;
                color: #ffffff;
            }
            QComboBox QAbstractItemView::item:selected {
                background-color: #dddddd;
                color: #000000;
            }
            QCheckBox {
                background: none;
                color: #ffffff;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border: 1px solid #555;
                background-color: #3C3F41;
            }
            QCheckBox::indicator:checked {
                image: url(":/qt-project.org/styles/commonstyle/images/check_on.png");
                background-color: #3C3F41;
                border: 1px solid #999;
            }
            QFrame {
                border: none;
            }
            QFrame#dragDropFrame {
                border: 2px dashed #aaa;
                border-radius: 5px;
            }
        """)

    def on_load_pdf_clicked(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Vyber PDF", "", "PDF Files (*.pdf)")
        if file_path:
            self.load_pdf_file(file_path)

    def load_pdf_file(self, pdf_path: str):
        fields = extract_form_fields(pdf_path)
        data = parse_personal_data(fields)
        self.jmeno_line.setText(data.get("jmeno", ""))
        self.narozeni_line.setText(data.get("narozeni", ""))
        self.adresa_line.setText(data.get("adresa", ""))
        self.psc_line.setText(data.get("psc", ""))

    def collect_factors(self):
        faktor_values = {}
        for i in range(6):
            f = self.faktor_combos[i].currentText()
            k = self.kategorie_combos[i].currentText()
            faktor_values[f"faktor{i+1}"] = f
            faktor_values[f"kategorie{i+1}"] = k
        return faktor_values

    def on_generate_docx(self):
        jmeno = self.jmeno_line.text()
        prijmeni = get_surname(jmeno)
        default_name = f"LP_{prijmeni}.docx"
        docx_path, _ = QFileDialog.getSaveFileName(self, "Uložit DOCX jako", default_name, "Docx Files (*.docx)")
        if not docx_path:
            return

        narozeni = self.narozeni_line.text()
        adresa = self.adresa_line.text()
        psc = self.psc_line.text()
        pozice = self.pozice_line.text()
        pobocka = self.pobocka_combo.currentText()
        konzultant = self.konzultant_combo.currentText()

        if self.remember_checkbox.isChecked():
            self.settings.setValue("consultant", konzultant)

        selected_pozice_popis = self.pozice_popis_combo.currentText()
        corresponding_popis_pozice = position_descriptions.get(selected_pozice_popis, "")

        generate_docx(
            jmeno=jmeno,
            narozeni=narozeni,
            adresa=adresa,
            psc=psc,
            pozice=pozice,
            pobocka=pobocka,
            konzultant=konzultant,
            pozice_popis=selected_pozice_popis,
            popis_pozice=corresponding_popis_pozice,
            faktor_values=self.collect_factors(),
            output_file=docx_path
        )
        QMessageBox.information(self, "Hotovo", f"Soubor {docx_path} byl vytvořen.")
        open_file(docx_path)

    def on_generate_pdf(self):
        jmeno = self.jmeno_line.text()
        prijmeni = get_surname(jmeno)
        default_name = f"LP_{prijmeni}.pdf"
        pdf_path, _ = QFileDialog.getSaveFileName(self, "Uložit PDF jako", default_name, "PDF Files (*.pdf)")
        if not pdf_path:
            return

        temp_docx = os.path.join(tempfile.gettempdir(), "temp_vysledny_posudek.docx")

        narozeni = self.narozeni_line.text()
        adresa = self.adresa_line.text()
        psc = self.psc_line.text()
        pozice = self.pozice_line.text()
        pobocka = self.pobocka_combo.currentText()
        konzultant = self.konzultant_combo.currentText()

        if self.remember_checkbox.isChecked():
            self.settings.setValue("consultant", konzultant)

        selected_pozice_popis = self.pozice_popis_combo.currentText()
        corresponding_popis_pozice = position_descriptions.get(selected_pozice_popis, "")

        generate_docx(
            jmeno=jmeno,
            narozeni=narozeni,
            adresa=adresa,
            psc=psc,
            pozice=pozice,
            pobocka=pobocka,
            konzultant=konzultant,
            pozice_popis=selected_pozice_popis,
            popis_pozice=corresponding_popis_pozice,
            faktor_values=self.collect_factors(),
            output_file=temp_docx
        )
        # Použijeme PDF generaci přes COM
        generate_pdf_with_comtypes(docx_file=temp_docx, pdf_file=pdf_path)
        if os.path.exists(temp_docx):
            os.remove(temp_docx)

        QMessageBox.information(self, "Hotovo", f"Soubor {pdf_path} byl vytvořen.")
        open_file(pdf_path)

def main():
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='dark_teal.xml', extra={'density_scale': '-1'})

    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
