# patch_checksum.py
import pefile
import sys
import os

def set_pe_checksum(file_path):
    """
    Öffnet eine PE-Datei (z.B. eine .exe), berechnet die korrekte Prüfsumme
    und schreibt diese in den Header der Datei.

    Args:
        file_path (str): Der Pfad zur zu patchenden .exe-Datei.
    """
    if not os.path.exists(file_path):
        print(f"FEHLER: Datei nicht gefunden unter '{file_path}'")
        return

    print(f"-> Öffne Datei: {os.path.basename(file_path)}")
    try:
        pe = pefile.PE(file_path)
        
        # Berechnet die korrekte Prüfsumme...
        new_checksum = pe.generate_checksum()
        
        # ...und setzt sie im Optional Header der Datei.
        pe.OPTIONAL_HEADER.CheckSum = new_checksum
        
        print("-> Schreibe neuen, validen Checksum...")
        # Die Änderungen werden in die Datei zurückgeschrieben.
        pe.write(file_path)
        print("-> Checksum erfolgreich gesetzt!")

    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}")

if __name__ == '__main__':
    # Das Skript erwartet den Pfad zur .exe als Argument in der Kommandozeile.
    if len(sys.argv) < 2:
        print("Verwendung: python patch_checksum.py <Pfad zur .exe-Datei>")
    else:
        exe_path = sys.argv[1]
        set_pe_checksum(exe_path)