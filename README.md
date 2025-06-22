# Ersatzteilkatalog-Generator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## Übersicht

**Ersatzteilkatalog-Generator** ist ein flexibles Open-Source-Tool zur automatisierten Erstellung von Ersatzteilkatalogen aus Stücklisten-Exceldateien (z.B. aus Autodesk Inventor, SolidWorks oder anderen CAD-Systemen).  
Das Tool unterstützt den Import, die flexible Spaltenzuordnung, Regel- und Mapping-Bearbeitung sowie die Ausgabe als Word-Dokument inklusive Inhaltsverzeichnis, Bildern und konfigurierbaren Layouts.

---

## Hauptfunktionen

- **Import mehrerer Stücklisten-Exceldateien** (auch in Ordnerstruktur, inkl. Unterbaugruppen)
- **Individuelle Spaltenzuordnung & Mapping** direkt im Editor
- **Rule-Engine** für automatische Spaltenmanipulation und Setzregeln
- **Bilderzuordnung** pro Position möglich
- **Automatische Generierung eines vollständigen Ersatzteilkatalogs** (Word .docx)
- **Benutzerfreundliches GUI** (Qt/PySide6), alle Konfigurationen direkt im Tool editierbar
- **Projekt-Management:** Projekte speichern/öffnen, eigene Vorlagen hinterlegen
- **Automatische Backups**
- **Open Source & MIT-Lizenz**

---

## Standard-Importaufbau (Excel)

**Typische Struktur einer Importdatei:**  
- **Zeile 1–4:** Metadaten (Titel, Zeichnungsnummer, Ersteller, Verwendung, Erstellungsdatum, Kunde etc.)
- **Zeile 5:** Spaltenüberschriften der Stücklistentabelle (Header)
- **Ab Zeile 6:** Stücklisten-Positionen (eigentliche Bauteile/Teile)

**Beispiel:**

| Objek | ANZAHL | Einheitenmenge | BEN (Benennung) | ZBEN (Zusatzbenennung) | NORM (Norm/Halbzeug) | Abmessung | Material | Masse | ZEICHNR (DB-Zeichnungsnummer) | Hersteller | Hersteller-Nr | Teileart | OSchutz | Klassifizierung | AFPS | Änderungsindex | Kommentar |
|-------|--------|----------------|-----------------|-----------------------|---------------------|-----------|----------|-------|-------------------------------|------------|---------------|----------|---------|----------------|------|----------------|-----------|
| 1001  | 2      | Stk            | Flanschplatte   | -                     | EN 1092-1           | 120x12    | S235JR   | 1.23  | ZN-12345                      | Muster AG  | FLP-1001      | Teil     | -       | Normteil       | -    | 0              | -         |
| 1002  | 12     | Stk            | Schraube M10x50 | -                     | DIN 933             |           | 8.8      | 0.04  |                               |            |               | Teil     | -       | C-Teil         | -    | 0              |           |

> **Hinweis:**  
> Die Spaltennamen und Reihenfolge können in der Firma variieren, sind aber im Tool per Editor flexibel zuordenbar (Mapping).  
> Der eigentliche Import startet **immer ab Zeile 6** (da Zeile 1-4 Metadaten und Zeile 5 Header sind).

**Meta-Header-Beispiel (Zeile 1–4, frei angeordnet):**

|            |           |           |                        |           |             |           |                |              |             |             |           |         |           |
|------------|-----------|-----------|------------------------|-----------|-------------|-----------|----------------|--------------|-------------|-------------|-----------|---------|-----------|
| Metadaten  | Titel:    | Titel2    | Zeichnungsnr. (STK):   | 37.120-1  | Verwendung: | Verwendung|                | Ersteller:   | M.Kohtz     | Kunden Zeichnungsnr. (STK): |           |         |           |
|            | Zusatzbenennung: | Untertitel2 | Zeichnungsnr. (Zchg): | 37.120-10 |           |           |                | Erstellungsdatum: | 12.04.2023 | Kunden Zeichnungsnr. (Zchg): |           |         |           |
| ...        | ...       | ...       | ...                    | ...       | ...         | ...       | ...            | ...          | ...         | ...         | ...       | ...     | ...       |

---

## Kurzanleitung

1. **Projekt anlegen oder öffnen**  
   → Bestehende Projekte laden oder neues anlegen.

2. **Stücklisten importieren**  
   → Importordner wählen (z.B. Export aus CAD-System).

3. **Spalten-Mapping & Setzregeln bearbeiten**  
   → Im Editor die relevanten Spalten zuordnen und Setzregeln anpassen.

4. **Optional: Bilder zu Positionen zuweisen**  
   → Grafiken können im Editor pro Position zugeordnet werden.

5. **Katalog generieren**  
   → Mit einem Klick wird das Word-Dokument erzeugt (inkl. Inhaltsverzeichnis und Bildintegration).

---

## Konfiguration & Erweiterung

- **Regeln und Spaltenzuordnung:**  
  Über den Editor kannst du Setzregeln, Mappings und Projektparameter jederzeit anpassen und speichern.
- **Eigene Vorlagen:**  
  Word-Vorlagen im Projekt hinterlegen – für individuelles Layout, Logo, CI etc.
- **Backups:**  
  Das Tool legt automatisch Sicherungen der Projektdatei an.

---

## Fehlerbehandlung & Support

- **Logs:**  
  Fehler werden im Logfile gespeichert (im Projektordner).
- **Typische Fehler:**  
  - Falsche/fehlende Spalten im Import → Mapping im Editor prüfen
  - Word nicht installiert → MS Word für DOCX-Ausgabe nötig (nur Windows)
- **Bug melden:**  
  [GitHub-Issues](https://github.com/leratos/Ersatzteilkatalog-Generator/issues) nutzen

---

## FAQ

**Kann ich beliebige Stücklisten importieren?**  
Ja, solange eine Exceldatei mit klaren Spalten vorliegt. Das Mapping ist flexibel.

**Funktioniert das Tool auf macOS/Linux?**  
Das GUI und die Datenlogik laufen überall, aber die Word-Integration (Docx-Export mit Layout) benötigt MS Word unter Windows.

**Kann ich eigene Regeln/Mappings speichern?**  
Ja, alles wird projektbasiert gesichert.

**Was mache ich, wenn das Importformat sich ändert?**  
Die Mapping- und Regel-Engine im Editor erlaubt es, Anpassungen vorzunehmen – bei komplexeren Änderungen bitte ein neues Mapping abspeichern oder [Issue einreichen](https://github.com/leratos/Ersatzteilkatalog-Generator/issues).

---

## Lizenz

MIT License (siehe [LICENSE](LICENSE))

---

## Project Summary (English)

**Ersatzteilkatalog-Generator** is a German-language open-source tool for the automated generation of spare parts catalogs from BOM Excel files (including CAD/ERP exports).  
Features include flexible column mapping, a rule engine for field processing, image assignment, and Word catalog export.  
Maintained by Marcus Kohtz ([leratos](https://github.com/leratos)), released under the MIT License.

---

*Maintainer: Marcus Kohtz (Signz-vision.de / GitHub: leratos)*
