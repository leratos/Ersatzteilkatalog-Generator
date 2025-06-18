# Ersatzteilkatalog-Generator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## Projektbeschreibung

**Ersatzteilkatalog-Generator** ist ein Open-Source-Tool für die strukturierte Erstellung von Ersatzteilkatalogen aus Stücklisten, wie sie typischerweise aus Autodesk Inventor exportiert werden.  
Die Anwendung ermöglicht das Einlesen, Filtern und hierarchische Verknüpfen von Stücklisten sowie die automatisierte Erstellung eines fertigen Word-Katalogs inkl. Inhaltsverzeichnis und Grafikzuordnung.

**Funktionen:**
- Einlesen mehrerer Stücklisten-Exceldateien (.xlsm/.xlsx)
- Flexible Spalten- und Headerzuordnung über Editor
- Verwaltung komplexer Baugruppenstrukturen (auch mit Unterbaugruppen)
- Benutzerfreundliche GUI (PySide6/Qt)
- Automatische Generierung von Ersatzteilkatalogen als Word-Dokument (.docx), inkl. Inhaltsverzeichnis und Seitennavigation
- Integration von Grafiken auf Positionsebene

---

## Maintainer und Verantwortlicher

**Projektleiter / Maintainer:**  
Marcus Kohtz (Signz-vision.de)  
Marcus.Kohtz@signz-vision.com

Die Open-Source-Version dieses Tools wird federführend von Marcus Kohtz entwickelt und betreut.  
Der Bezug zum Entwickler und Inhaber wird im Git-Log und im Projekt-Header explizit dokumentiert.

---

## Anforderungen

- Python 3.10+  
- Benötigte Libraries:  
  - pandas, openpyxl, python-docx, PySide6, pillow, pywin32 (Windows)

Siehe `requirements.txt` für die vollständige Paketliste.

---

## Installation & Start

```bash
pip install -r requirements.txt
python main.py
