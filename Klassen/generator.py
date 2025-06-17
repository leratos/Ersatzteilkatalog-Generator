from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import datetime
import os
import sys
import re
from PIL import Image

class DocxGenerator:
    def __init__(self, data: dict, main_bom, author_name: str, template_path: str, output_path: str, auto_update_fields: bool, project_path: str):
        self.data = data; self.main_bom = main_bom; self.author_name = author_name
        self.template_path = template_path; self.output_path = output_path
        self.auto_update_fields = auto_update_fields
        self.project_path = project_path
        self.doc = Document(self.template_path)

    def run(self) -> bool:
        try:
            self._update_header_footer()
            self._create_cover_sheet()
            self._create_toc()
            self._create_assembly_section(self.data)
            self._replace_graphic_placeholders()
            self.doc.save(self.output_path)
            
            if self.auto_update_fields:
                self._update_fields_and_refs_with_word()

            return True
        except Exception as e:
            print(f"FEHLER bei der Dokumenterstellung: {e}")
            return False

    def _update_fields_and_refs_with_word(self):
        """
        Starts Word, updates the TOC, parses the TOC's text content with a robust regex
        to get page numbers, and then replaces the placeholders.
        """
        print("Starte erweiterte Feld-Aktualisierung mit MS Word...")
        if sys.platform != "win32":
            print("WARNUNG: Automatische Aktualisierung ist nur auf Windows möglich.")
            return

        word = None 
        try:
            import win32com.client as win32
            
            abs_path = os.path.abspath(self.output_path)
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(abs_path)
            
            print("  -> Aktualisiere Inhaltsverzeichnis...")
            if doc.TablesOfContents.Count > 0:
                doc.TablesOfContents(1).Update()
            
            page_map = {}
            print("  -> Lese Seitenzahlen aus Inhaltsverzeichnis-Text...")
            if doc.TablesOfContents.Count > 0:
                toc_text = doc.TablesOfContents(1).Range.Text
                
                # --- FINALE, ROBUSTE LOGIK: Regex, die den Kern der ZN extrahiert ---
                # Sucht nach dem Muster XX.XXX-XX... und ignoriert alles danach
                pattern = re.compile(r'\((\d{2}\.\d{3}-\d{1,3}(?:\.\d{1,3})?)(?:.*?)\)[\s\S]*?\t(\d+)')

                matches = pattern.finditer(toc_text)
                for match in matches:
                    znr = match.group(1).replace(' ', '')
                    page = match.group(2)
                    page_map[znr] = page
            
            print(f"  -> Gefundenes Seiten-Mapping: {len(page_map)} Einträge: {page_map}")
            if not page_map and doc.TablesOfContents.Count > 0:
                print("    [WARNUNG] Konnte keine Seitenzahlen aus dem Inhaltsverzeichnis extrahieren. Überprüfen Sie das Format.")

            if page_map:
                print("  -> Ersetze Seiten-Referenzen im Dokument (manuelle Methode)...")
                for table in doc.Tables:
                    for row in table.Rows:
                        for cell in row.Cells:
                            if "[REF:" in cell.Range.Text:
                                cell_text = cell.Range.Text.strip()
                                match = re.search(r'\[REF:(.*?)\]', cell_text)
                                if match:
                                    znr_in_ref = match.group(1)
                                    if znr_in_ref in page_map:
                                        page = page_map[znr_in_ref]
                                        cell.Range.Text = f"S. {page}"

            doc.Close(SaveChanges=True)
            word.Quit()
            word = None
            print("Erweiterte Feld-Aktualisierung erfolgreich abgeschlossen.")

        except ImportError:
            print("FEHLER: Die Bibliothek 'pywin32' wird für die automatische Aktualisierung benötigt.")
        except Exception as e:
            print(f"FEHLER bei der automatischen Aktualisierung: {e}")
        finally:
            if word is not None:
                word.Quit()

    def _create_table_for_assembly(self, items):
        """Creates table and inserts page reference PLACEHOLDERS."""
        headers = ['Pos.', 'Menge', 'Benennung', 'Bestellnummer', 'Information', 'Seite']
        table = self.doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        
        hdr_cells = table.rows[0].cells
        for i, header_text in enumerate(headers): 
            cell = hdr_cells[i]
            cell.text = header_text
            cell.paragraphs[0].runs[0].font.bold = True
        
        tr = table.rows[0]._tr
        trPr = tr.get_or_add_trPr(); tblHeader = OxmlElement('w:tblHeader'); trPr.append(tblHeader)
        
        def get_pos_key(item):
            try: return float(item.get('POS', 0))
            except (ValueError, TypeError): return float('inf')
        sorted_items = sorted(items, key=get_pos_key)

        for i, item in enumerate(sorted_items):
            row_cells = table.add_row().cells
            row_cells[0].text = f"{item.get('POS', ''):g}"
            row_cells[1].text = item.get('Menge', '')
            self._add_multiline_text(row_cells[2], item.get('Benennung_Formatiert', ''))
            row_cells[3].text = item.get('Bestellnummer_Kunde', '')
            self._add_multiline_text(row_cells[4], item.get('Information', ''))
            
            if item.get('is_assembly'):
                teilenummer_raw = str(item.get('Teilenummer', ''))
                clean_znr = teilenummer_raw.split('(')[0].strip().replace(' ', '')
                row_cells[5].text = f"[REF:{clean_znr}]"
            else:
                row_cells[5].text = ""

            if (i % 2) != 0:
                for cell in row_cells: self._set_cell_shading(cell, "DAE9F8")

        col_widths = [Cm(1.2), Cm(2.0), Cm(5.1), Cm(3.8), Cm(3.8), Cm(1.3)]
        for i, width in enumerate(col_widths):
            for cell in table.columns[i].cells: cell.width = width

    # Der Rest der Klasse bleibt unverändert.
    def _create_cover_sheet(self):
        p_title = self.doc.add_paragraph(); p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_title = p_title.add_run(self.main_bom.titel or "[TITEL]"); font_title = run_title.font; font_title.name = 'Calibri'; font_title.size = Cm(1.5); font_title.bold = True
        p_line = self.doc.add_paragraph(); p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER; p_line.add_run("_________________________________________________")
        p_subject = self.doc.add_paragraph(); p_subject.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_subject = p_subject.add_run("Ersatzteilkatalog"); run_subject.font.size = Cm(0.8)
        p_grafik = self.doc.add_paragraph(); p_grafik.alignment = WD_ALIGN_PARAGRAPH.CENTER
        image_path = None
        for ext in ['.png', '.jpg']:
            path = os.path.join(self.project_path, "Grafik", f"EL{ext}")
            if os.path.exists(path): image_path = path; break
        if image_path: self._add_scaled_picture(p_grafik, image_path, Cm(16), Cm(15))
        self.doc.add_page_break()
    def _create_toc(self):
        self.doc.add_heading("Inhaltsverzeichnis", level=1)
        paragraph = self.doc.add_paragraph(); run = paragraph.add_run()
        fldChar_begin = OxmlElement('w:fldChar'); fldChar_begin.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve'); instrText.text = ' TOC \\o "1-3" \\h \\z \\u '
        fldChar_end = OxmlElement('w:fldChar'); fldChar_end.set(qn('w:fldCharType'), 'end')
        run._r.append(fldChar_begin); run._r.append(instrText); run._r.append(fldChar_end)
        self.doc.add_page_break()
    def _add_scaled_picture(self, paragraph, image_path, max_width, max_height):
        try:
            img = Image.open(image_path)
            original_width_px, original_height_px = img.size
            if original_width_px == 0: return
            aspect_ratio = float(original_height_px) / float(original_width_px)
            height_at_max_width = max_width * aspect_ratio
            if height_at_max_width > max_height: paragraph.add_run().add_picture(image_path, height=max_height)
            else: paragraph.add_run().add_picture(image_path, width=max_width)
        except Exception as e:
            print(f"Bildfehler: {e}"); paragraph.add_run(f"[Bild {os.path.basename(image_path)} nicht ladbar]").italic = True
    def _replace_graphic_placeholders(self):
        grafik_folder = os.path.join(self.project_path, "Grafik"); max_breite = Cm(16); max_hoehe = Cm(20)
        for p in self.doc.paragraphs:
            if "[GRAFIK_PLATZHALTER_" in p.text:
                znr = p.text.split('_')[-1].replace(']', ''); image_path = None
                for ext in ['.png', '.jpg', '.jpeg']:
                    path = os.path.join(grafik_folder, f"{znr}{ext}")
                    if os.path.exists(path): image_path = path; break
                p.clear() 
                if image_path: self._add_scaled_picture(p, image_path, max_breite, max_hoehe)
                else: p.add_run("[Keine Grafik zugeordnet]").italic = True
    def _update_header_footer(self):
        title = self.main_bom.titel or ""; zusatz = self.main_bom.zusatzbenennung or ""; full_title = f"{title} - {zusatz}".strip(' -')
        zeich_nr = self.main_bom.kundennummer or self.main_bom.zeichnungsnummer or ""; full_zeich_nr = f"{zeich_nr} (EL)"
        verwend = self.main_bom.verwendung or ""; creation_date = datetime.date.today().strftime('%d.%m.%Y')
        replacements = {'[TITEL]': full_title, '[THEMA]': "Ersatzteilkatalog", '[ZEICH]': full_zeich_nr, '[VERWEND]': verwend, '[AUTOR]': self.author_name, '[EDATUM]': creation_date}
        for section in self.doc.sections:
            for part in [section.header, section.footer, section.first_page_header, section.first_page_footer]:
                if part is not None:
                    for table in part.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    for run in p.runs:
                                        for key, value in replacements.items():
                                            if key in run.text: run.text = run.text.replace(key, str(value))
                    for p in part.paragraphs: 
                        for run in p.runs:
                            for key, value in replacements.items():
                                if key in run.text: run.text = run.text.replace(key, str(value))
    def _create_assembly_section(self, assembly_data):
        if not assembly_data: return
        title = f"{assembly_data.get('Benennung', '')} ({assembly_data.get('Teilenummer', '')})"; heading = self.doc.add_heading(title, level=1)
        heading.paragraph_format.keep_with_next = True
        styles = self.doc.styles
        if 'Überschrift 1' in styles: heading.style = styles['Überschrift 1']
        elif 'Heading 1' in styles: heading.style = styles['Heading 1']
        p_grafik = self.doc.add_paragraph(); run = p_grafik.add_run(); run.text = f"[GRAFIK_PLATZHALTER_{assembly_data.get('Teilenummer')}]"; p_grafik.paragraph_format.keep_with_next = True
        self.doc.add_paragraph() 
        if assembly_data.get('children'): self._create_table_for_assembly(assembly_data.get('children'))
        self.doc.add_page_break()
        for child in assembly_data.get('children', []):
            if child.get('is_assembly'): self._create_assembly_section(child)
    def _add_multiline_text(self, cell, text):
        cell.text = ''; p = cell.paragraphs[0]; lines = str(text).split('\n')
        for i, line in enumerate(lines):
            if i > 0: p.add_run().add_break()
            p.add_run(line)
    def _set_cell_shading(self, cell, fill_color):
        tc_pr = cell._tc.get_or_add_tcPr(); shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), fill_color); tc_pr.append(shd)

