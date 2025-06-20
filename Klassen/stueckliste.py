# -*- coding: utf-8 -*-
"""
Dieses Modul definiert die Logik zum Einlesen und Verarbeiten der Stücklisten.

Es enthält zwei Hauptklassen:
- Stueckliste: Repräsentiert eine einzelne Stückliste als Objekt.
- BomProcessor: Dient als "Manager", der einen ganzen Ordner von Stücklisten-
  Dateien verarbeitet, die Logik der RuleEngine anwendet und die
  Baugruppenstruktur erstellt.
"""

import os
import pandas as pd
import openpyxl

# NEUER IMPORT: Wir importieren unsere neue RuleEngine.
from Klassen.rule_engine import RuleEngine


class Stueckliste:
    """
    Repräsentiert eine einzelne Stückliste mit ihren Metadaten und den
    ROHDATEN ihrer Positionen. Die Datenaufbereitung erfolgt im BomProcessor.
    """
    def __init__(self, filepath: str, config_manager):
        """
        Initialisiert ein Stueckliste-Objekt.

        Args:
            filepath (str): Der vollständige Pfad zur Excel-Datei.
            config_manager: Der ConfigManager für die Konfiguration.
        """
        self.filepath = filepath
        self.config = config_manager
        self.titel = None
        self.zusatzbenennung = None
        self.zeichnungsnummer = None
        self.kundennummer = None
        self.verwendung = None
        self.positionen = []
        self.is_loaded = False
        self._load_data()

    def _load_data(self):
        """
        Private Methode, die die Excel-Datei öffnet, die Daten extrahiert,
        filtert und als Rohdaten speichert. Die Logik zur Aufbereitung
        wurde in den BomProcessor verlagert.
        """
        print(f"  [DEBUG] Lese Datei: {os.path.basename(self.filepath)}")
        try:
            # Phase 1: Header-Daten auslesen (bleibt gleich)
            col_indices = self.config.get_column_indices()
            workbook = openpyxl.load_workbook(self.filepath, data_only=True)
            if 'Import' not in workbook.sheetnames:
                print("    [WARNUNG] Tabellenblatt 'Import' nicht gefunden.")
                return

            sheet = workbook['Import']
            self.titel = sheet[self.config.get_header_cell('titel')].value
            raw_znr = str(sheet[self.config.get_header_cell('zeichnungsnummer')].value or '')
            self.zeichnungsnummer = raw_znr.split('(')[0].strip().replace(' ', '')
            self.zusatzbenennung = sheet[self.config.get_header_cell('zusatzbenennung')].value
            self.kundennummer = sheet[self.config.get_header_cell('kundennummer')].value
            self.verwendung = sheet[self.config.get_header_cell('verwendung')].value

            # Phase 2: Positionsdaten mit Pandas einlesen (bleibt gleich)
            df_items = pd.read_excel(
                self.filepath, sheet_name='Import', header=None, skiprows=5
            )
            df_items.dropna(how='all', inplace=True)

            # Phase 3: Filtern der Positionsdaten (bleibt gleich)
            pos_col = col_indices.get('POS', 0)
            type_col = col_indices.get('Teileart', 12)
            if df_items.shape[1] <= max(pos_col, type_col):
                print("    [WARNUNG] Nicht genügend Spalten für die Verarbeitung.")
                return

            df_items = df_items[pd.to_numeric(df_items.iloc[:, pos_col], errors='coerce').notna()]
            df_items = df_items[df_items.iloc[:, pos_col] == df_items.iloc[:, pos_col].astype(int)]
            df_items = df_items[df_items.iloc[:, type_col].isin([1, 4, 5])]
            print(f"    [DEBUG] Nach Filterung: {len(df_items)} gültige Positionen.")

            # Phase 4: Rohdaten extrahieren
            processed_positions = []
            for index, row in df_items.iterrows():
                pos_dict = {}
                for name, col_idx in col_indices.items():
                    pos_dict[name] = row.get(col_idx)
                processed_positions.append(pos_dict)

            if not processed_positions:
                self.is_loaded = True
                return

            df_final = pd.DataFrame(processed_positions).fillna('')

            self.positionen = df_final.to_dict(orient='records')
            self.is_loaded = True
            print(f"    [DEBUG] Erfolgreich {len(self.positionen)} Roh-Positionen geladen.")
        except Exception as e:
            print(f"FEHLER bei der Verarbeitung von '{self.filepath}': {e}")
            self.is_loaded = False

    def __repr__(self):
        """Gibt eine menschenlesbare Repräsentation des Objekts zurück."""
        return f"Stueckliste(ZN: '{self.zeichnungsnummer}', Titel: '{self.titel}')"


class BomProcessor:
    """Verwaltet einen Stapel von Stücklisten-Objekten."""
    def __init__(self, folder_path: str, config_manager):
        """
        Initialisiert den Prozessor.
        """
        self.folder_path = folder_path
        self.boms = {}
        self.config = config_manager

    def run(self):
        """Führt den gesamten Prozess aus: Einlesen, Regeln anwenden und Verknüpfen."""
        self._load_all_boms()
        self._apply_generation_rules()  # KORREKTER AUFRUF
        self._link_assemblies()
        return self.boms

    def _load_all_boms(self):
        """Lädt alle .xlsm/.xlsx-Dateien im Ordner als Stueckliste-Objekte."""
        print("\n--- Starte Einlese-Prozess im Ordner ---")
        for filename in sorted(os.listdir(self.folder_path)):
            if filename.lower().endswith((".xlsm", ".xlsx")) and not filename.startswith('~'):
                filepath = os.path.join(self.folder_path, filename)
                bom = Stueckliste(filepath, config_manager=self.config)
                if bom.is_loaded and bom.zeichnungsnummer:
                    self.boms[bom.zeichnungsnummer] = bom

    def _apply_generation_rules(self):
        """
        Wendet die Setzregeln auf alle geladenen Positionen an.
        """
        print("\n--- Starte Anwendung der Setzregeln ---")
        rules = self.config.config.get("generation_rules", {})
        if not rules:
            print("  [INFO] Keine Setzregeln definiert. Überspringe...")
            return

        rule_engine = RuleEngine(rules)
        processed_count = 0

        for bom in self.boms.values():
            if not bom.positionen:
                continue

            df = pd.DataFrame(bom.positionen)

            generated_data_df = df.apply(
                rule_engine.process_row, axis=1, result_type='expand'
            )

            df = pd.concat([df, generated_data_df], axis=1)

            def format_menge(row):
                menge_val = row.get('Menge_val')
                einheit_raw = str(row.get('Einheit', ''))
                einheit = 'Stk' if einheit_raw in ['1', '1.0'] else einheit_raw
                return f"{menge_val:g} {einheit}".strip() if pd.notna(menge_val) and menge_val != '' else ""

            df['Menge'] = df.apply(format_menge, axis=1)

            bom.positionen = df.to_dict(orient='records')
            processed_count += len(bom.positionen)
        
        print(f"--- Regeln auf {processed_count} Positionen angewendet. ---")

    def _link_assemblies(self):
        """Verknüpft die geladenen Stücklisten zu einer Baugruppen-Hierarchie."""
        print("\n--- Starte Verknüpfung der Baugruppen ---")
        link_count = 0
        for bom in self.boms.values():
            for position in bom.positionen:
                teilenummer_raw = position.get('Teilenummer')
                if not teilenummer_raw:
                    continue

                teilenummer_str = str(teilenummer_raw).strip()
                cleaned_teilenummer = teilenummer_str.split('(')[0].strip()
                final_teilenummer = cleaned_teilenummer.replace(' ', '')

                if final_teilenummer and final_teilenummer in self.boms:
                    position['sub_assembly'] = self.boms[final_teilenummer]
                    print(f"  [LINK] Pos '{position.get('POS')}' in '{bom.zeichnungsnummer}' -> Baugruppe '{final_teilenummer}'")
                    link_count += 1
        print(f"--- Verknüpfung abgeschlossen. {link_count} Baugruppen verknüpft. ---")
