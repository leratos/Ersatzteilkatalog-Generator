import pandas as pd
import openpyxl # Wird von pandas zum Lesen von .xlsx/.xlsm benötigt

def parse_bom_excel(filepath: str) -> dict | None:
    """
    Liest eine einzelne Inventor-Stücklisten-Excel-Datei (.xlsm) ein,
    extrahiert Header-Daten sowie die Positionsliste anhand fixer Spaltenpositionen
    und filtert diese.

    Args:
        filepath: Der Dateipfad zur .xlsm-Datei.

    Returns:
        Ein Dictionary mit den extrahierten Daten oder None bei einem Fehler.
    """
    try:
        # --- 1. Header-Daten mit openpyxl auslesen (bleibt gleich) ---
        workbook = openpyxl.load_workbook(filepath, data_only=True)
        if 'Import' not in workbook.sheetnames:
            print(f"FEHLER: Das Tabellenblatt 'Import' wurde in der Datei '{filepath}' nicht gefunden.")
            return None
        sheet = workbook['Import']
        titel = sheet['C2'].value
        zeichnungsnummer = sheet['E2'].value

        # --- 2. Positionsliste mit pandas einlesen (positionsbasiert) ---
        # Wir überspringen die ersten 6 Zeilen und lesen OHNE Header,
        # sodass pandas die Spalten mit 0, 1, 2, ... indiziert.
        df_items = pd.read_excel(filepath, sheet_name='Import', header=None, skiprows=5)

        # Definition der Spaltenpositionen (0-basiert, A=0, B=1, etc.)
        COL_POS = {
            'Objekt': 0,
            'Teileart': 12
        }

        # --- 3. Daten filtern und bereinigen ---
        df_items.dropna(how='all', inplace=True)

        # Filter 1: Behalte nur Positionen, deren 'Objekt' (Spalte A) eine ganze Zahl ist.
        # Wir greifen über den Index .iloc[:, COL_POS['Objekt']] auf die Spalte zu
        df_items['Objekt_numeric'] = pd.to_numeric(df_items.iloc[:, COL_POS['Objekt']], errors='coerce')
        df_items.dropna(subset=['Objekt_numeric'], inplace=True)
        df_items = df_items[df_items['Objekt_numeric'] == df_items['Objekt_numeric'].astype(int)]

        # Filter 2: Behalte nur Teilearten 1, 4, 5 (Spalte M)
        valid_part_types = [1, 4, 5]
        df_items = df_items[df_items.iloc[:, COL_POS['Teileart']].isin(valid_part_types)]

        # --- 4. Relevante Spalten auswählen und benennen ---
        # Key = Neuer, sauberer Name | Value = Spaltenindex (0-basiert)
        spalten_mapping = {
            'POS': 0,               # Spalte A
            'Menge': 1,             # Spalte B
            'Benennung': 3,         # Spalte D
            'Zusatzbenennung': 4,   # Spalte E
            'Abmessung': 6,         # Spalte G
            'Teilenummer': 9,       # Spalte J
            'Hersteller': 10,       # Spalte K
            'Hersteller_Nr': 11     # Spalte L
        }

        # Wähle die benötigten Spalten über ihre Positionsnummern aus
        df_final = df_items.iloc[:, list(spalten_mapping.values())]
        # Weise den ausgewählten Spalten die neuen, sauberen Namen zu
        df_final.columns = list(spalten_mapping.keys())
        
        positions_liste = df_final.to_dict(orient='records')

        # --- 5. Ergebnis zusammenstellen ---
        return {
            'titel': titel,
            'zeichnungsnummer': zeichnungsnummer,
            'positionen': positions_liste
        }

    except FileNotFoundError:
        print(f"FEHLER: Die Datei '{filepath}' wurde nicht gefunden.")
        return None
    except Exception as e:
        print(f"FEHLER: Ein unerwarteter Fehler ist aufgetreten bei der Verarbeitung von '{filepath}'. Details: {e}")
        return None


# --- Beispiel für die Verwendung ---
if __name__ == '__main__':
    # ERSETZEN SIE DIESEN DATEINAMEN mit dem Namen Ihrer echten Test-Stückliste.
    dateiname = 'IHRE_STUECKLISTE.xlsm'
    
    bom_data = parse_bom_excel(dateiname)

    if bom_data:
        print("--- Erfolgreich eingelesene Daten ---")
        print(f"Titel: {bom_data['titel']}")
        print(f"Zeichnungsnummer: {bom_data['zeichnungsnummer']}")
        print("\n--- Gefilterte Positionen ---")
        
        import json
        print(json.dumps(bom_data['positionen'], indent=2, ensure_ascii=False))

