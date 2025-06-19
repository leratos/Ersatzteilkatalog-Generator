# -*- coding: utf-8 -*-
"""
Dieses Modul definiert die Logik zum Einlesen und Verarbeiten der Stücklisten.

Es enthält zwei Hauptklassen:
- Stueckliste: Repräsentiert eine einzelne Stückliste als Objekt. Sie ist
  verantwortlich für das Öffnen einer Excel-Datei, das Extrahieren der
  relevanten Daten aus den korrekten Zellen und Spalten und das Anwenden
  von Geschäftslogik zur Datenbereinigung und -aufbereitung.
- BomProcessor: Dient als "Manager", der einen ganzen Ordner von Stücklisten-
  Dateien verarbeitet. Er erstellt für jede Datei ein Stueckliste-Objekt
  und verknüpft diese anschließend miteinander, um die hierarchische
  Baugruppenstruktur zu erstellen.

Autor: Marcus Kohtz (Signz-vision.de)
Zuletzt geändert: 13.06.2025
"""

import os
import pandas as pd
import openpyxl


class Stueckliste:
    """
    Repräsentiert eine einzelne Stückliste mit ihren Metadaten und Positionen.

    Jedes Objekt dieser Klasse entspricht einer physischen Excel-Datei. Die Klasse
    ist dafür verantwortlich, die Daten aus der Datei zu laden und in eine
    strukturierte, programminterne Form zu überführen.
    """
    def __init__(self, filepath: str, config_manager):
        """
        Initialisiert ein Stueckliste-Objekt.

        Args:
            filepath (str): Der vollständige Pfad zur Excel-Datei (.xlsm/.xlsx).
            config_manager: Der ConfigManager für die Konfiguration.
        """
        self.filepath = filepath
        self.config = config_manager
        # --- Instanzvariablen für die Metadaten der Stückliste ---
        self.titel = None
        self.zusatzbenennung = None
        self.zeichnungsnummer = None  # Die bereinigte Haupt-ZN der Baugruppe
        self.kundennummer = None
        self.verwendung = None

        # --- Instanzvariablen für die Verarbeitung ---
        self.positionen = []  # Liste der Positions-Dictionaries
        self.is_loaded = False  # Flag, um erfolgreiches Laden zu markieren

        # Startet den Ladevorgang direkt bei der Objekterstellung.
        self._load_data()

    def _load_data(self):
        """
        Private Methode, die die Excel-Datei öffnet, die Daten extrahiert,
        filtert und aufbereitet. Dies ist die Kernlogik des Parsers.
        """
        print(f"  [DEBUG] Lese Datei: {os.path.basename(self.filepath)}")
        try:
            # --- Phase 1: Header-Daten auslesen ---
            col_indices = self.config.get_column_indices()

            workbook = openpyxl.load_workbook(self.filepath, data_only=True)
            if 'Import' not in workbook.sheetnames:
                print("    [WARNUNG] Tabellenblatt 'Import' nicht gefunden.")
                return

            sheet = workbook['Import']
            self.titel = sheet[self.config.get_header_cell('titel')].value

            # Bereinige die Haupt-Zeichnungsnummer direkt beim Einlesen, um
            # Konsistenz im gesamten Programm sicherzustellen.
            raw_znr = str(sheet[self.config.get_header_cell('zeichnungsnummer')].value or '')
            self.zeichnungsnummer = raw_znr.split('(')[0].strip().replace(' ', '')
            
            self.zusatzbenennung = sheet[self.config.get_header_cell('zusatzbenennung')].value
            self.kundennummer = sheet[self.config.get_header_cell('kundennummer')].value
            self.verwendung = sheet[self.config.get_header_cell('verwendung')].value

            # --- Phase 2: Positionsdaten mit Pandas einlesen ---
            df_items = pd.read_excel(
                self.filepath, sheet_name='Import', header=None, skiprows=5
            )
            df_items.dropna(how='all', inplace=True)
            print(f"    [DEBUG] Rohdaten geladen: {len(df_items)} Zeilen")

            pos_col = col_indices['POS']
            type_col = col_indices['Teileart']

            # --- Phase 3: Filtern der Positionsdaten ---
            # Stellt sicher, dass die Datei die Mindestanzahl an Spalten hat.
            if df_items.shape[1] <= max(pos_col, type_col):
                print("    [WARNUNG] Nicht genügend Spalten für die Verarbeitung.")
                return

            # Filterlogik, um nur relevante Zeilen zu behalten.
            df_items = df_items[pd.to_numeric(df_items.iloc[:, 0], errors='coerce').notna()]
            df_items = df_items[df_items.iloc[:, 0] == df_items.iloc[:, 0].astype(int)]
            df_items = df_items[df_items.iloc[:, 12].isin([1, 4, 5])]
            print(f"    [DEBUG] Nach Filterung: {len(df_items)} gültige Positionen.")

            # --- Phase 4: Daten extrahieren und aufbereiten ---
            # Definiere, welche Spalte welchen internen Namen bekommen soll.
            
            # Konvertiere die gefilterten Reihen in eine Liste von Dictionaries.
            # Dieser Ansatz ist robuster als der direkte Zugriff auf Spalten.
            processed_positions = []
            for index, row in df_items.iterrows():
                pos_dict = {}
                for name, col_idx in col_indices.items():
                    pos_dict[name] = row.get(col_idx, '')
                processed_positions.append(pos_dict)

            if not processed_positions:
                self.is_loaded = True
                return

            df_final = pd.DataFrame(processed_positions).fillna('')

            # --- Phase 5: Geschäftslogik anwenden ---
            # Hier werden die Rohdaten in das Format gebracht, das der Katalog benötigt.
            
            def format_benennung_multiline(row):
                b = str(row.get('Benennung', '')).strip()
                z = str(row.get('Zusatzbenennung', '')).strip()
                return f"{b}\n{z}" if z else b

            def format_menge(row):
                m = row['Menge_val']
                e_raw = str(row.get('Einheit', ''))
                e = 'Stk' if e_raw in ['1', '1.0'] else e_raw
                return f"{m:g} {e}".strip() if pd.notna(m) and m != '' else ""

            def get_bestellnummer(row):
                raw_teilenr = str(row.get('Teilenummer', ''))
                clean_teilenr = raw_teilenr.split('(')[0].strip()
                return str(row.get('AFPS') or clean_teilenr or row.get('Hersteller_Nr') or '')

            def get_information(row):
                has_internal_number = bool(str(row.get('AFPS') or '').strip()) or \
                                      bool(str(row.get('Teilenummer') or '').strip())
                if has_internal_number:
                    return ""
                else:
                    n = f"{str(row.get('Norm',''))} {str(row.get('Abmessung',''))}".strip()
                    h = f"{str(row.get('Hersteller',''))} / {str(row.get('Hersteller_Nr',''))}".strip(' /')
                    return f"{n}\n{h}".strip()

            df_final['Benennung_Formatiert'] = df_final.apply(format_benennung_multiline, axis=1)
            df_final['Menge'] = df_final.apply(format_menge, axis=1)
            df_final['Bestellnummer_Kunde'] = df_final.apply(get_bestellnummer, axis=1)
            df_final['Information'] = df_final.apply(get_information, axis=1)
            
            output_columns = self.config.config.get("output_columns", [])
            for col_config in output_columns:
                source_id = col_config.get("source_id")
                col_id = col_config.get("id")
                # Wenn keine Datenquelle -> manuelle Spalte. Füge sie zum DataFrame hinzu.
                if not source_id and col_id not in df_final.columns:
                    df_final[col_id] = ""

            self.positionen = df_final.to_dict(orient='records')
            self.is_loaded = True
            print(f"    [DEBUG] Erfolgreich {len(self.positionen)} Positionen verarbeitet.")
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

        Args:
            folder_path (str): Der Pfad zum Ordner, der die Excel-Dateien enthält.
            config_manager: Der ConfigManager für die Konfiguration.
        """
        self.folder_path = folder_path
        self.boms = {}  # Dictionary, um ZN auf Stueckliste-Objekte zu mappen
        self.config = config_manager # Speichere den ConfigManager

    def run(self):
        """Führt den gesamten Prozess aus: Einlesen und Verknüpfen."""
        self._load_all_boms()
        self._link_assemblies()
        return self.boms

    def _load_all_boms(self):
        """Lädt alle .xlsm/.xlsx-Dateien im Ordner als Stueckliste-Objekte."""
        print("\n--- Starte Einlese-Prozess im Ordner ---")
        # Sortiert die Dateiliste für eine konsistente Verarbeitungsreihenfolge.
        for filename in sorted(os.listdir(self.folder_path)):
            # Ignoriert temporäre Excel-Dateien, die mit '~' beginnen.
            if filename.lower().endswith((".xlsm", ".xlsx")) and not filename.startswith('~'):
                filepath = os.path.join(self.folder_path, filename)
                bom = Stueckliste(filepath, config_manager=self.config)
                # Füge das Objekt nur hinzu, wenn es erfolgreich geladen wurde
                # und eine gültige Zeichnungsnummer hat.
                if bom.is_loaded and bom.zeichnungsnummer:
                    self.boms[bom.zeichnungsnummer] = bom
    
    def _link_assemblies(self):
        """Verknüpft die geladenen Stücklisten zu einer Baugruppen-Hierarchie."""
        print("\n--- Starte Verknüpfung der Baugruppen ---")
        link_count = 0
        for bom in self.boms.values():
            for position in bom.positionen:
                teilenummer_raw = position.get('Teilenummer')
                if not teilenummer_raw:
                    continue
                
                # Bereinige die Teilenummer, um eine Übereinstimmung zu finden.
                teilenummer_str = str(teilenummer_raw).strip()
                cleaned_teilenummer = teilenummer_str.split('(')[0].strip()
                final_teilenummer = cleaned_teilenummer.replace(' ', '')
                
                # Wenn die bereinigte Nummer im Dictionary der geladenen Stücklisten
                # existiert, ist es eine Unterbaugruppe.
                if final_teilenummer and final_teilenummer in self.boms: 
                    # Füge einen direkten Verweis auf das untergeordnete Objekt hinzu.
                    position['sub_assembly'] = self.boms[final_teilenummer]
                    print(f"  [LINK] Pos '{position.get('POS')}' in '{bom.zeichnungsnummer}' -> Baugruppe '{final_teilenummer}'")
                    link_count += 1
        print(f"--- Verknüpfung abgeschlossen. {link_count} Baugruppen verknüpft. ---")