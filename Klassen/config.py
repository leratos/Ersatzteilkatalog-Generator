# -*- coding: utf-8 -*-
"""
Dieses Modul definiert den ConfigManager.

Die Klasse ist verantwortlich für das Laden, Verwalten und Speichern der
Projekt-Konfiguration (mapping.json). Sie stellt die Brücke zwischen den
flexiblen Einstellungen des Benutzers und der festen Logik des Programms dar.
"""

import os
import json
from openpyxl.utils import column_index_from_string

class ConfigManager:
    """Verwaltet die Lese- und Schreibvorgänge für die mapping.json."""

    def __init__(self, project_path: str):
        """
        Initialisiert den Manager für einen spezifischen Projektordner.

        Args:
            project_path (str): Der Pfad zum aktuellen Projekt.
        """
        self.config_path = os.path.join(project_path, 'mapping.json')
        self._default_config = self._get_default_config()
        self.config = self._load_config()

    def _get_default_config(self):
        """Definiert die Standard-Konfiguration als Fallback."""
        return {
            "header_mapping": {
                "titel": "D2",
                "zeichnungsnummer": "G2",
                "zusatzbenennung": "D3",
                "kundennummer": "N3",
                "verwendung": "J2"
            },
            "column_mapping": {
                "POS": "A",
                "Menge_val": "B",
                "Einheit": "C",
                "Benennung": "D",
                "Zusatzbenennung": "E",
                "Norm": "F",
                "Abmessung": "G",
                "Teilenummer": "J",
                "Hersteller": "K",
                "Hersteller_Nr": "L",
                "AFPS": "P",
                "Teileart": "M"
            },
            "output_columns": [
                {"id": "std_pos", "header": "Pos.", "width_percent": 8, "source_id": "POS", "type": "standard"},
                {"id": "std_menge", "header": "Menge", "width_percent": 10, "source_id": "Menge", "type": "standard"},
                {"id": "std_benennung", "header": "Benennung", "width_percent": 30, "source_id": "Benennung_Formatiert", "type": "standard"},
                {"id": "std_bestellnr", "header": "Bestellnummer", "width_percent": 22, "source_id": "Bestellnummer_Kunde", "type": "standard"},
                {"id": "std_info", "header": "Information", "width_percent": 22, "source_id": "Information", "type": "standard"},
                {"id": "std_seite", "header": "Seite", "width_percent": 8, "source_id": "Seite", "type": "standard"},
            ],
            "generation_rules": {
                # Standardregel für die Benennung: Kombiniert zwei Felder mit Zeilenumbruch
                "Benennung_Formatiert": {
                    "type": "combine",
                    "sources": ["Benennung", "Zusatzbenennung"],
                    "separator": "\\n"
                },
                # Standardregel für die Bestellnummer: Nimmt das erste verfügbare Feld
                "Bestellnummer_Kunde": {
                    "type": "prioritized_list",
                    "sources": ["AFPS", "Teilenummer", "Hersteller_Nr"]
                }

            },
            "table_styles": {
              "base_style": "Table Grid",
              "header_bold": True,
              "header_font_color": "FFFFFF",
              "header_shading_color": "4F81BD",
              "shading_enabled": True,
              "shading_color": "DAE9F8"
            },
            "formatting_options": {
                "assembly_title_format": "{benennung} ({teilenummer})",
                "add_space_after_title": True,
                "table_on_new_page": False,
                "blank_pages_before_toc": 0,
                "cover_sheet_type": "default",
                "cover_sheet_path": "",
                "blank_pages_type": "blank",
                "blank_pages_path": "",
                "toc_on_new_page": True
            }

        }

    def get_all_available_data_ids(self) -> list:
        """
        Gibt eine Liste aller möglichen Daten-IDs zurück, die für die Ausgabe
        verwendet werden können.
        """
        input_fields = list(
            self.config.get("column_mapping", {}).keys()
            )
        generated_fields = [
            "Benennung_Formatiert", "Menge", 
            "Bestellnummer_Kunde", "Information", "Seite"
        ]
        all_fields = sorted(list(set(input_fields + generated_fields)))
        generated_fields = list(self.config.get("generation_rules", {}).keys())
        manual_fields = ["Seite"]

        all_fields = sorted(list(set(input_fields + generated_fields + manual_fields)))
        return [""] + all_fields

    def _load_config(self):
        """
        Lädt die Konfiguration aus der mapping.json.
        Wenn die Datei nicht existiert, wird sie mit Standardwerten erstellt.
        """
        if not os.path.exists(self.config_path):
            print(f"INFO: 'mapping.json' nicht gefunden. Erstelle neue Datei mit Standardwerten.")
            self._create_default_config()
        
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:

                loaded_config = json.load(f)
            is_dirty = False
            # Stelle sicher, dass alle Haupt-Keys vorhanden sind
            for key, value in self._default_config.items():
                if key not in loaded_config:
                    loaded_config[key] = value
                    is_dirty = True
            # Speichere die ggf. ergänzte Konfiguration zurück
                elif isinstance(value, dict):
                    for sub_key, sub_value in value.items():
                        if sub_key not in loaded_config.get(key, {}):
                            loaded_config[key][sub_key] = sub_value
                            is_dirty = True

            if is_dirty:
                self.save_config(loaded_config)
            return loaded_config
        except (json.JSONDecodeError, IOError) as e:
            print(f"FEHLER: Konnte 'mapping.json' nicht laden. Verwende Standardwerte. Fehler: {e}")
            return self._default_config

    def _create_default_config(self):
        self.save_config(self._default_config)

    def get_header_cell(self, key: str) -> str:
        """Gibt die Zelle für einen Header-Wert zurück."""
        return self.config.get("header_mapping", {}).get(key)

    def get_column_map(self) -> dict:
        """Gibt das Mapping von internem Namen zu Excel-Spaltenbuchstabe zurück."""
        return self.config.get(
            "column_mapping", self._default_config["column_mapping"]
        )

    def get_column_indices(self) -> dict:
        """
        Konvertiert die Excel-Spaltenbuchstaben in numerische, nullbasierte Indizes
        für die Verwendung mit der pandas-Bibliothek.
        """
        column_map = self.config.get("column_mapping", {})
        index_map = {}
        for name, letter in column_map.items():
            if not letter:
                continue
            try:
                # Konvertiert 'A' -> 1, 'B' -> 2, etc. und zieht 1 ab für 0-basiert.
                index_map[name] = column_index_from_string(letter) - 1
            except ValueError:
                print(f"WARNUNG: Ungültiger Spaltenbuchstabe '{letter}' in der Konfiguration für '{name}'.")
        return index_map
    
    def save_config(self, new_config: dict):
        """Speichert die übergebene Konfiguration in die mapping.json."""
        self.config = new_config
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            print("INFO: Konfiguration erfolgreich in 'mapping.json' gespeichert.")
        except IOError as e:
            print(f"FEHLER: Konnte 'mapping.json' nicht speichern. Fehler: {e}")