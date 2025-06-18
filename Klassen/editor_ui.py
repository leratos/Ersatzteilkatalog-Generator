# -*- coding: utf-8 -*-
"""
Dieses Modul definiert das Konfigurationsfenster (Editor).

Es ermöglicht dem Benutzer, die Zuordnung von Excel-Spalten zu den
internen Datenfeldern des Programms dynamisch zu ändern, ohne den
Quellcode bearbeiten zu müssen.
"""

from PySide6 import QtWidgets, QtCore
import openpyxl
from openpyxl.utils import get_column_letter

class ConfigEditorWindow(QtWidgets.QDialog):
    """
    Ein Dialogfenster zur Bearbeitung der Projekt-Konfiguration (mapping.json).
    """
    def __init__(self, config_manager, parent=None):
        """
        Initialisiert das Editor-Fenster.

        Args:
            config_manager (ConfigManager): Die Instanz des ConfigManagers,
                                            die die Konfiguration verwaltet.
            parent (QWidget, optional): Das übergeordnete Widget. Defaults to None.
        """
        super().__init__(parent)
        self.config_manager = config_manager
        self.setWindowTitle("Editor")
        self.setMinimumSize(800, 600)

        self.layout = QtWidgets.QVBoxLayout(self)
        self.tabs = QtWidgets.QTabWidget()
        self.layout.addWidget(self.tabs)
        
        self.mapping_tab = QtWidgets.QWidget()
        self.layout_tab = QtWidgets.QWidget()

        self.tabs.addTab(self.mapping_tab, "Spaltenzuordnung (Import)")
        self.tabs.addTab(self.layout_tab, "Katalog-Layout (Export)")

        self._setup_mapping_tab()
        self._setup_layout_tab()

        self.button_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Save | 
            QtWidgets.QDialogButtonBox.StandardButton.Cancel
        )
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.layout.addWidget(self.button_box)

    def _setup_mapping_tab(self):
        """Erstellt die Widgets für den "Spaltenzuordnung"-Tab."""
        layout = QtWidgets.QVBoxLayout(self.mapping_tab)
        
        self.load_sample_button = QtWidgets.QPushButton("Muster-Stückliste laden...")
        self.load_sample_button.clicked.connect(self._load_sample_bom)
        layout.addWidget(self.load_sample_button)

        header_group = QtWidgets.QGroupBox("Header-Felder (Zellen)")
        header_layout = QtWidgets.QGridLayout()
        header_group.setLayout(header_layout)
        
        self.header_widgets = {}
        header_config = self.config_manager.config.get("header_mapping", {})
        for row, (key, value) in enumerate(header_config.items()):
            label = QtWidgets.QLabel(f"{key}:")
            line_edit = QtWidgets.QLineEdit(value)
            self.header_widgets[key] = line_edit
            header_layout.addWidget(label, row, 0)
            header_layout.addWidget(line_edit, row, 1)

        self.column_group = QtWidgets.QGroupBox("Positions-Felder (Spalten)")
        self.column_layout = QtWidgets.QGridLayout()
        self.column_group.setLayout(self.column_layout)

        self.column_widgets = {}
        self.available_columns = [""] # Leere Option am Anfang
        self._create_column_comboboxes()
            
        layout.addWidget(header_group)
        layout.addWidget(self.column_group)
        layout.addStretch()

    def _setup_layout_tab(self):
        """Erstellt die UI für den Katalog-Layout-Editor."""
        layout = QtWidgets.QGridLayout(self.layout_tab)
        
        self.available_fields_list = QtWidgets.QListWidget()
        self.selected_columns_list = QtWidgets.QListWidget()
        
        # Fülle die verfügbaren Felder. Diese sind intern definiert.
        available_ids = ["POS", "Menge", "Benennung_Formatiert", "Bestellnummer_Kunde", "Information", "Seite"]
        self.available_fields_list.addItems(available_ids)

        # Fülle die ausgewählten Spalten aus der aktuellen Konfiguration.
        output_config = self.config_manager.config.get("output_columns", [])
        for col in output_config:
            self.selected_columns_list.addItem(col.get("header"))

        # Knöpfe zur Steuerung
        add_button = QtWidgets.QPushButton("->"); remove_button = QtWidgets.QPushButton("<-")
        up_button = QtWidgets.QPushButton("Hoch"); down_button = QtWidgets.QPushButton("Runter")
        
        # Layout für die mittleren Knöpfe
        button_layout = QtWidgets.QVBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(add_button); button_layout.addWidget(remove_button)
        button_layout.addSpacing(40)
        button_layout.addWidget(up_button); button_layout.addWidget(down_button)
        button_layout.addStretch()

        layout.addWidget(QtWidgets.QLabel("Verfügbare Datenfelder"), 0, 0)
        layout.addWidget(self.available_fields_list, 1, 0)
        layout.addLayout(button_layout, 1, 1)
        layout.addWidget(QtWidgets.QLabel("Spalten im Katalog"), 0, 2)
        layout.addWidget(self.selected_columns_list, 1, 2)
        
        # Verbinde Signale mit Funktionen
        add_button.clicked.connect(self._add_column)
        remove_button.clicked.connect(self._remove_column)
        up_button.clicked.connect(self._move_column_up)
        down_button.clicked.connect(self._move_column_down)

    def _add_column(self):
        """Fügt ein ausgewähltes Feld zur Liste der Katalog-Spalten hinzu."""
        selected_item = self.available_fields_list.currentItem()
        if selected_item:
            # Füge den Header-Text zur UI-Liste hinzu.
            self.selected_columns_list.addItem(selected_item.text())

    def _remove_column(self):
        """Entfernt eine Spalte aus der Katalog-Liste."""
        selected_item = self.selected_columns_list.currentItem()
        if selected_item:
            self.selected_columns_list.takeItem(self.selected_columns_list.row(selected_item))

    def _move_column_up(self):
        """Bewegt die ausgewählte Spalte eine Position nach oben."""
        current_row = self.selected_columns_list.currentRow()
        if current_row > 0:
            item = self.selected_columns_list.takeItem(current_row)
            self.selected_columns_list.insertItem(current_row - 1, item)
            self.selected_columns_list.setCurrentRow(current_row - 1)

    def _move_column_down(self):
        """Bewegt die ausgewählte Spalte eine Position nach unten."""
        current_row = self.selected_columns_list.currentRow()
        if current_row < self.selected_columns_list.count() - 1:
            item = self.selected_columns_list.takeItem(current_row)
            self.selected_columns_list.insertItem(current_row + 1, item)
            self.selected_columns_list.setCurrentRow(current_row + 1)

    def _create_column_comboboxes(self):
        """(Neu) Erstellt oder aktualisiert die Dropdown-Menüs für die Spalten."""
        while self.column_layout.count():
            child = self.column_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        self.column_widgets = {}
        column_config = self.config_manager.config.get("column_mapping", {})
        
        for row, (key, current_value) in enumerate(column_config.items()):
            label = QtWidgets.QLabel(f"{key}:")
            combo_box = QtWidgets.QComboBox()
            combo_box.addItems(self.available_columns)
            
            full_text_to_set = ""
            for option in self.available_columns:
                if option.startswith(current_value + ' -'):
                    full_text_to_set = option
                    break
            
            if full_text_to_set:
                combo_box.setCurrentText(full_text_to_set)
            else:
                # Fallback, falls der Buchstabe nicht in den Optionen ist
                combo_box.addItem(current_value)
                combo_box.setCurrentText(current_value)

            self.column_widgets[key] = combo_box
            self.column_layout.addWidget(label, row, 0)
            self.column_layout.addWidget(combo_box, row, 1)

    def _load_sample_bom(self):
        """Öffnet eine Excel-Datei und liest die Spaltenüberschriften aus Zeile 5."""
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Muster-Stückliste auswählen", "", "Excel-Dateien (*.xlsm *.xlsx)"
        )
        if not file_path:
            return

        try:
            workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            sheet = workbook['Import']
            
            headers = []
            # Lese die Spalten aus Zeile 5 (dort stehen die Header)
            for cell in sheet[5]:
                if cell.value:
                    col_letter = get_column_letter(cell.column)
                    # Speichere Buchstabe und Wert, z.B. "A - Objekt"
                    headers.append(f"{col_letter} - {cell.value}")
                
            if headers:
                # --- KORRIGIERTE LOGIK: Speichere den vollen Text ---
                self.available_columns = [""] + headers
                self._create_column_comboboxes()
                QtWidgets.QMessageBox.information(self, "Erfolg", f"{len(headers)} Spalten aus der Musterdatei geladen.")
            else:
                QtWidgets.QMessageBox.warning(self, "Fehler", "Konnte keine Spaltenüberschriften in Zeile 5 finden.")

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Fehler beim Lesen der Datei", str(e))

    def accept(self):
        """Wird aufgerufen, wenn der Benutzer auf "Speichern" klickt."""
        new_header_mapping = {}
        for key, widget in self.header_widgets.items():
            new_header_mapping[key] = widget.text().upper()
            
        new_column_mapping = {}
        for key, widget in self.column_widgets.items():
            # Lese den ausgewählten Text aus der ComboBox
            # und nehme nur den Spaltenbuchstaben am Anfang.
            selected_text = widget.currentText()
            new_column_mapping[key] = selected_text.split(' - ')[0]

        new_output_columns = []
        for i in range(self.selected_columns_list.count()):
            header_text = self.selected_columns_list.item(i).text()
            known_ids = ["POS", "Menge", "Benennung_Formatiert", "Bestellnummer_Kunde", "Information", "Seite"]
            data_id = header_text if header_text not in known_ids else header_text # Vereinfachung
            
            # Verwende Standardbreiten, die später anpassbar gemacht werden können
            new_output_columns.append({"id": data_id, "header": header_text, "width_cm": 2.5})

        new_config = self.config_manager.config
        new_config["header_mapping"] = new_header_mapping
        new_config["column_mapping"] = new_column_mapping
        new_config["output_columns"] = new_output_columns
        
        self.config_manager.save_config(new_config)
        
        super().accept()