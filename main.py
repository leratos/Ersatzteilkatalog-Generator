# -*- coding: utf-8 -*-
"""
Haupt-Startpunkt für den Ersatzteilkatalog-Generator.

Dieses Skript ist verantwortlich für das Starten der grafischen Benutzeroberfläche (GUI)
und die initiale Auswahl des Projektordners durch den Benutzer. Es dient als
minimalistischer Einstiegspunkt, der die gesamte Anwendungslogik an das
MainWindow-Objekt delegiert.

Autor: Marcus Kohtz (Signz-vision.de)
Zuletzt geändert: 13.06.2025
"""

import sys
import os
from PySide6 import QtWidgets
# Importiert die Hauptklasse der Benutzeroberfläche aus dem Klassen-Verzeichnis.
from Klassen.ui import MainWindow


def main():
    """
    Die Hauptfunktion, die die gesamte Anwendung initialisiert und startet.

    Ablauf:
    1. Erstellt eine QApplication-Instanz, die für jede PySide-Anwendung notwendig ist.
    2. Öffnet einen nativen Datei-Dialog, damit der Benutzer einen Ordner für sein
       Projekt auswählen oder erstellen kann.
    3. Wenn kein Ordner ausgewählt wird (Abbruch), wird die Anwendung beendet.
    4. Erstellt eine Instanz des Hauptfensters (MainWindow) und übergibt den
       ausgewählten Projektpfad.
    5. Zeigt das Fenster an und startet die Event-Loop der Anwendung, die auf
       Benutzerinteraktionen wartet.
    """
    # Initialisiert die Anwendungsumgebung.
    app = QtWidgets.QApplication(sys.argv)

    # Öffnet einen Dialog, um einen existierenden Ordner auszuwählen.
    # Dies stellt sicher, dass alle relevanten Dateien (Stücklisten, Grafiken,
    # Speicherdateien) an einem zentralen Ort gebündelt sind.
    project_path = QtWidgets.QFileDialog.getExistingDirectory(
        None,
        "Wählen oder erstellen Sie einen Projektordner",
        # Startet den Dialog standardmäßig im Benutzerverzeichnis für
        # eine bessere Benutzererfahrung.
        os.path.expanduser("~")
    )

    # Wenn der Benutzer den Dialog schließt, ohne einen Ordner auszuwählen,
    # soll die Anwendung sauber beendet werden.
    if not project_path:
        sys.exit()

    # Erstellt das Hauptfenster und übergibt den Projektpfad, damit die
    # Klasse weiß, wo sie arbeiten soll.
    window = MainWindow(project_path=project_path)
    window.show()

    # Startet die Anwendung und wartet auf Ereignisse (Klicks, etc.).
    # sys.exit sorgt für ein sauberes Beenden.
    sys.exit(app.exec())


# Dieser Standard-Python-Block stellt sicher, dass die main()-Funktion nur
# dann ausgeführt wird, wenn dieses Skript direkt gestartet wird (und nicht,
# wenn es von einem anderen Skript importiert wird).
if __name__ == '__main__':
    main()
