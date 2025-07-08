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
import shutil
from PySide6 import QtCore, QtGui, QtWidgets
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
    if not check_and_prepare_template():
        sys.exit(1)  # Beende das Programm, wenn keine Vorlage vorhanden ist.

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
    pixmap_path = "logo.png"
    if os.path.exists(pixmap_path):
        pixmap = QtGui.QPixmap(pixmap_path)
    else:
        pixmap = QtGui.QPixmap()
        
    splash = QtWidgets.QSplashScreen(pixmap)
    splash.showMessage(
        "Initialisiere Projekt und lade Stücklisten...\nDies kann einen Moment dauern.",
        QtCore.Qt.AlignmentFlag.AlignBottom | QtCore.Qt.AlignmentFlag.AlignCenter,
        QtCore.Qt.black
    )
    splash.show()
    app.processEvents()

    window = MainWindow(project_path=project_path)
    window.show()
    splash.finish(window)

    # Startet die Anwendung und wartet auf Ereignisse (Klicks, etc.).
    # sys.exit sorgt für ein sauberes Beenden.
    sys.exit(app.exec())

def check_and_prepare_template() -> bool:
    """
    Prüft, ob der 'Vorlagen'-Ordner und eine Dokumentvorlage existieren.
    Wenn nicht, wird der Benutzer aufgefordert, eine Vorlage auszuwählen.

    Returns:
        bool: True, wenn eine Vorlage vorhanden oder erfolgreich erstellt wurde,
              sonst False.
    """
    template_folder = "Vorlagen"
    template_path_docx = os.path.join(template_folder, "DOK-Vorlage.docx")
    template_path_docm = os.path.join(template_folder, "DOK-Vorlage.docm")

    # Prüfen, ob eine der beiden Vorlagen-Dateien existiert.
    if os.path.exists(template_path_docx) or os.path.exists(template_path_docm):
        return True

    # Wenn nicht, informiere den Benutzer und starte den Setup-Prozess.
    os.makedirs(template_folder, exist_ok=True)

    msg_box = QtWidgets.QMessageBox()
    msg_box.setIcon(QtWidgets.QMessageBox.Icon.Information)
    msg_box.setWindowTitle("Master-Vorlage fehlt")
    msg_box.setText(
        "Die Master-Dokumentvorlage wurde nicht gefunden.\n\n"
        "Bitte wählen Sie im nächsten Schritt Ihre .docx- oder .docm-Datei aus, "
        "die als Vorlage für alle zukünftigen Kataloge dienen soll."
    )
    msg_box.setStandardButtons(QtWidgets.QMessageBox.StandardButton.Ok)
    msg_box.exec()

    file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
        None,
        "Master-Vorlage auswählen",
        "",
        "Word-Dokumente (*.docx *.docm)",
    )

    if not file_path:
        QtWidgets.QMessageBox.critical(
            None, "Abbruch", "Ohne eine Master-Vorlage kann das Programm nicht starten."
        )
        return False

    # Kopiere und benenne die ausgewählte Datei korrekt um.
    try:
        if file_path.lower().endswith(".docm"):
            destination_path = template_path_docm
        else:
            destination_path = template_path_docx
        
        shutil.copy(file_path, destination_path)
        QtWidgets.QMessageBox.information(
            None, "Erfolg", f"Die Vorlage wurde erfolgreich nach '{destination_path}' kopiert."
        )
        return True
    except Exception as e:
        QtWidgets.QMessageBox.critical(
            None, "Fehler beim Kopieren", f"Die Vorlage konnte nicht kopiert werden:\n{e}"
        )
        return False
    
# Dieser Standard-Python-Block stellt sicher, dass die main()-Funktion nur
# dann ausgeführt wird, wenn dieses Skript direkt gestartet wird (und nicht,
# wenn es von einem anderen Skript importiert wird).
if __name__ == '__main__':
    main()
