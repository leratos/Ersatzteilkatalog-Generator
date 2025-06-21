# -*- coding: utf-8 -*-
"""
Dieses Modul definiert den DocxGenerator.

Die Klasse ist verantwortlich für die Erstellung des finalen Word-Katalogs
basierend auf den aufbereiteten Daten und den Konfigurationen.
"""

import datetime
import os
import re
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from PIL import Image


class DocxGenerator:
    """Erstellt das Word-Dokument für den Ersatzteilkatalog."""

    def __init__(
        self,
        data: dict,
        main_bom,
        author_name: str,
        template_path: str,
        output_path: str,
        auto_update_fields: bool,
        project_path: str,
        config_manager,
    ):
        self.data = data
        self.main_bom = main_bom
        self.author_name = author_name
        self.template_path = template_path
        self.output_path = output_path
        self.auto_update_fields = auto_update_fields
        self.project_path = project_path
        self.config = config_manager
        self.doc = Document(self.template_path)

    def run(self) -> bool:
        """Führt den gesamten Prozess der Dokumenterstellung aus."""
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
        Startet Word, aktualisiert das Inhaltsverzeichnis und ersetzt die
        Seiten-Referenzen.
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

            # Das Dokument wird jetzt direkt geöffnet. Der "Säuberungs"-Schritt
            # ist für .docx-Dateien nicht notwendig.
            doc = word.Documents.Open(abs_path)
            
            print("  -> Aktualisiere Inhaltsverzeichnis...")
            if doc.TablesOfContents.Count > 0:
                doc.TablesOfContents(1).Update()

            page_map = {}
            print("  -> Lese Seitenzahlen aus Inhaltsverzeichnis-Text...")
            if doc.TablesOfContents.Count > 0:
                toc_text = doc.TablesOfContents(1).Range.Text
                pattern = re.compile(
                    r"\((\d{2}\.\d{3}-\d{1,3}(?:\.\d{1,3})?)(?:.*?)\)[\s\S]*?\t(\d+)"
                )
                for match in pattern.finditer(toc_text):
                    page_map[match.group(1).replace(" ", "")] = match.group(2)
            
            if not page_map and doc.TablesOfContents.Count > 0:
                print("    [WARNUNG] Konnte keine Seitenzahlen aus dem Inhaltsverzeichnis extrahieren.")

            if page_map:
                print("  -> Ersetze Seiten-Referenzen im Dokument...")
                for table in doc.Tables:
                    for row in table.Rows:
                        for cell in row.Cells:
                            match = re.search(r"\[REF:(.*?)\]", cell.Range.Text)
                            if match and match.group(1) in page_map:
                                cell.Range.Text = f"S. {page_map[match.group(1)]}"

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
        """Erstellt die Tabelle dynamisch basierend auf der Konfiguration."""
        output_columns = self.config.config.get("output_columns", [])
        if not output_columns:
            print("WARNUNG: Keine Ausgabespalten in der Konfiguration definiert.")
            return

        table_styles = self.config.config.get("table_styles", {})
        base_style = table_styles.get("base_style", "Table Grid")
        header_bold = table_styles.get("header_bold", True)
        shading_enabled = table_styles.get("shading_enabled", True)
        shading_color = table_styles.get("shading_color", "DAE9F8")

        headers = [col["header"] for col in output_columns]
        table = self.doc.add_table(rows=1, cols=len(headers))
        table.style = base_style
        table.width = Cm(16.5)
        table.autofit = False
        table.allow_autofit = False

        hdr_cells = table.rows[0].cells
        tr = table.rows[0]._tr
        trPr = tr.get_or_add_trPr()
        tblHeader = OxmlElement("w:tblHeader")
        trPr.append(tblHeader)
        for i, header_text in enumerate(headers):
            cell = hdr_cells[i]
            cell.text = header_text
            if header_bold:
                cell.paragraphs[0].runs[0].font.bold = True

        sorted_items = sorted(
            items,
            key=lambda item: (
                float(item.get("POS", 0))
                if str(item.get("POS", "0")).replace(".", "", 1).isdigit()
                else float("inf")
            ),
        )

        for i, item in enumerate(sorted_items):
            row_cells = table.add_row().cells
            for col_idx, col_config in enumerate(output_columns):
                source_id = col_config.get("source_id")
                col_id = col_config.get("id")
                key_for_data = source_id if source_id else col_id

                cell_text = ""
                if col_id == "std_seite":
                    if item.get("is_assembly", False) and item.get("Teilenummer"):
                        teilenummer_raw = str(item.get("Teilenummer", ""))
                        clean_znr = teilenummer_raw.split("(")[0].strip().replace(" ", "")
                        if clean_znr:
                            cell_text = f"[REF:{clean_znr}]"
                else:
                    value = item.get(key_for_data, "")
                    if key_for_data == "POS" and isinstance(value, (int, float)):
                        cell_text = f"{value:g}"
                    else:
                        cell_text = str(value)

                self._add_multiline_text(row_cells[col_idx], cell_text)

            if shading_enabled and (i % 2) != 0:
                for cell in row_cells:
                    self._set_cell_shading(cell, shading_color)

        total_percent = sum(c.get("width_percent", 10) for c in output_columns)
        if total_percent > 0:
            for i, col_config in enumerate(output_columns):
                percent = col_config.get("width_percent", 10)
                col_width = int(
                    table.width.cm * (percent / total_percent) * 360000
                )
                for row in table.rows:
                    row.cells[i].width = col_width
                table.columns[i].width = col_width

    def _create_cover_sheet(self):
        p_title = self.doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_title = p_title.add_run(self.main_bom.titel or "[TITEL]")
        font_title = run_title.font
        font_title.name = 'Calibri'
        font_title.size = Cm(1.5)
        font_title.bold = True
        
        p_line = self.doc.add_paragraph("_________________________________________________")
        p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        p_subject = self.doc.add_paragraph()
        p_subject.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_subject = p_subject.add_run("Ersatzteilkatalog")
        run_subject.font.size = Cm(0.8)
        
        p_grafik = self.doc.add_paragraph()
        p_grafik.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        image_path = next((os.path.join(self.project_path, "Grafik", f"EL{ext}") for ext in ['.png', '.jpg'] if os.path.exists(os.path.join(self.project_path, "Grafik", f"EL{ext}"))), None)
        if image_path:
            self._add_scaled_picture(p_grafik, image_path, Cm(16), Cm(15))
        
        self.doc.add_page_break()

    def _create_toc(self):
        self.doc.add_heading("Inhaltsverzeichnis", level=1)
        paragraph = self.doc.add_paragraph()
        run = paragraph.add_run()
        fldChar_begin = OxmlElement("w:fldChar"); fldChar_begin.set(qn("w:fldCharType"), "begin")
        instrText = OxmlElement("w:instrText"); instrText.set(qn("xml:space"), "preserve"); instrText.text = ' TOC \\o "1-3" \\h \\z \\u '
        fldChar_end = OxmlElement("w:fldChar"); fldChar_end.set(qn("w:fldCharType"), "end")
        run._r.append(fldChar_begin); run._r.append(instrText); run._r.append(fldChar_end)
        self.doc.add_page_break()

    def _add_scaled_picture(self, paragraph, image_path, max_width, max_height):
        try:
            with Image.open(image_path) as img:
                original_width_px, original_height_px = img.size
            if original_width_px == 0 or original_height_px == 0: return
            
            aspect_ratio = float(original_height_px) / float(original_width_px)
            if max_width * aspect_ratio > max_height:
                paragraph.add_run().add_picture(image_path, height=max_height)
            else:
                paragraph.add_run().add_picture(image_path, width=max_width)
        except Exception as e:
            print(f"Bildfehler: {e}")
            paragraph.add_run(f"[Bild {os.path.basename(image_path)} nicht ladbar]").italic = True

    def _replace_graphic_placeholders(self):
        grafik_folder = os.path.join(self.project_path, "Grafik")
        for p in self.doc.paragraphs:
            if "[GRAFIK_PLATZHALTER_" in p.text:
                znr = p.text.split('_')[-1].replace(']', '')
                image_path = next((os.path.join(grafik_folder, f"{znr}{ext}") for ext in ['.png', '.jpg', '.jpeg'] if os.path.exists(os.path.join(grafik_folder, f"{znr}{ext}"))), None)
                p.clear() 
                if image_path:
                    self._add_scaled_picture(p, image_path, Cm(16), Cm(20))
                else:
                    p.add_run("[Keine Grafik zugeordnet]").italic = True

    def _update_header_footer(self):
        full_title = f"{self.main_bom.titel or ''} - {self.main_bom.zusatzbenennung or ''}".strip(' -')
        full_zeich_nr = f"{self.main_bom.kundennummer or self.main_bom.zeichnungsnummer or ''} (EL)"
        creation_date = datetime.date.today().strftime('%d.%m.%Y')
        replacements = {'[TITEL]': full_title, '[THEMA]': "Ersatzteilkatalog", '[ZEICH]': full_zeich_nr, '[VERWEND]': self.main_bom.verwendung or "", '[AUTOR]': self.author_name, '[EDATUM]': creation_date}
        
        for section in self.doc.sections:
            for part in [section.header, section.footer, section.first_page_header, section.first_page_footer]:
                if part:
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
        title = f"{assembly_data.get('Benennung', '')} ({assembly_data.get('Teilenummer', '')})"
        heading = self.doc.add_heading(title, level=1)
        heading.paragraph_format.keep_with_next = True
        
        styles = self.doc.styles
        if 'Überschrift 1' in styles:
            heading.style = styles['Überschrift 1']
        elif 'Heading 1' in styles:
            heading.style = styles['Heading 1']

        p_grafik = self.doc.add_paragraph()
        p_grafik.add_run(f"[GRAFIK_PLATZHALTER_{assembly_data.get('Teilenummer')}]")
        p_grafik.paragraph_format.keep_with_next = True
        self.doc.add_paragraph()
        
        if assembly_data.get('children'):
            self._create_table_for_assembly(assembly_data.get('children'))
        
        self.doc.add_page_break()
        for child in assembly_data.get('children', []):
            if child.get('is_assembly'):
                self._create_assembly_section(child)

    def _add_multiline_text(self, cell, text):
        cell.text = ''
        p = cell.paragraphs[0]
        for i, line in enumerate(str(text).split('\n')):
            if i > 0: p.add_run().add_break()
            p.add_run(line)
            
    def _set_cell_shading(self, cell, fill_color):
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:fill'), fill_color)
        tc_pr.append(shd)
