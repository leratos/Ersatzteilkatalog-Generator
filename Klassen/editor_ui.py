# -*- coding: utf-8 -*-
"""
Dieses Modul definiert das UI-Fenster für den umfassenden Konfigurations-Editor.

Es ermöglicht dem Benutzer die Bearbeitung von:
1. Spaltenzuordnung (Import)
2. Katalog-Layout (Export)
3. Setzregeln (Daten-Generierung)
4. Tabellen-Design (Präsentation)
"""

import re
import time
from functools import partial

from PySide6 import QtCore, QtGui, QtWidgets



class ConfigEditorWindow(QtWidgets.QDialog):
    """
    Ein Dialogfenster zur Bearbeitung der kompletten Projekt-Konfiguration.
    """

    def __init__(self, config_manager, available_excel_columns, parent=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.available_columns = [""] + available_excel_columns
        self.current_rules = self.config_manager.config.get(
            "generation_rules", {}
        )
        self.current_target_field = None

        self.rule_type_map = {
            "Priorisierte Liste": "prioritized_list",
            "Werte kombinieren": "combine",
            "Bedingte Zuweisung": "conditional",
        }
        self.operator_map = {
            "ist gleich": "is",
            "ist nicht gleich": "is_not",
            "ist leer": "is_empty",
            "ist nicht leer": "is_not_empty",
            "enthält": "contains",
        }

        self.setWindowTitle("Konfigurations-Editor")
        self.setMinimumSize(950, 750)
        self._setup_ui()
        self._connect_signals()
        self._populate_target_fields_list()

        self._load_design_settings()


        if self.target_list.count() > 0:
            self.target_list.setCurrentRow(0)
        self._update_total_width()

    # --------------------------------------------------------------------------
    # --- UI Setup Methoden ---
    # --------------------------------------------------------------------------

    def _setup_ui(self):
        """Erstellt alle UI-Elemente und ordnet sie im Layout an."""
        self.layout = QtWidgets.QVBoxLayout(self)
        self.tabs = QtWidgets.QTabWidget()
        self.layout.addWidget(self.tabs)

        self.mapping_tab = QtWidgets.QWidget()
        self.layout_tab = QtWidgets.QWidget()
        self.rules_tab = QtWidgets.QWidget()

        self.design_tab = QtWidgets.QWidget()


        self.tabs.addTab(self.mapping_tab, "Spaltenzuordnung (Import)")
        self.tabs.addTab(self.layout_tab, "Katalog-Layout (Export)")
        self.tabs.addTab(self.rules_tab, "Setzregeln (Generierung)")

        self.tabs.addTab(self.design_tab, "Tabellen-Design")


        self._setup_mapping_tab()
        self._setup_layout_tab()
        self._setup_rules_tab()

        self._setup_design_tab()


        self.button_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Save
            | QtWidgets.QDialogButtonBox.StandardButton.Cancel
        )
        self.layout.addWidget(self.button_box)


    def _setup_design_tab(self):
        """Erstellt die UI für den 'Tabellen-Design'-Tab."""
        layout = QtWidgets.QVBoxLayout(self.design_tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)

        # --- Gruppe für Deckblatt ---
        cover_group = QtWidgets.QGroupBox("Deckblatt-Optionen")
        cover_layout = QtWidgets.QFormLayout(cover_group)
        self.cover_type_combo = QtWidgets.QComboBox()
        self.cover_type_combo.addItems(["Standard (Titel + Bild)", "Externe DOCX-Datei"])
        self.cover_path_input = QtWidgets.QLineEdit()
        self.cover_path_input.setReadOnly(True)
        cover_path_button = QtWidgets.QPushButton("...")
        cover_path_button.clicked.connect(lambda: self._select_external_doc(self.cover_path_input))
        cover_path_layout = QtWidgets.QHBoxLayout()
        cover_path_layout.addWidget(self.cover_path_input)
        cover_path_layout.addWidget(cover_path_button)
        cover_layout.addRow("Deckblatt-Typ:", self.cover_type_combo)
        cover_layout.addRow("Pfad zur DOCX:", cover_path_layout)
        layout.addWidget(cover_group)

        # --- Gruppe für zusätzliche Seiten ---
        pages_group = QtWidgets.QGroupBox("Zusätzliche Seiten (vor Inhaltsverzeichnis)")
        pages_layout = QtWidgets.QFormLayout(pages_group)
        self.pages_type_combo = QtWidgets.QComboBox()
        self.pages_type_combo.addItems(["Leere Seiten", "Externe DOCX-Datei"])
        self.blank_pages_spinbox = QtWidgets.QSpinBox()
        self.blank_pages_spinbox.setRange(0, 10)
        self.pages_path_input = QtWidgets.QLineEdit()
        self.pages_path_input.setReadOnly(True)
        pages_path_button = QtWidgets.QPushButton("...")
        pages_path_button.clicked.connect(lambda: self._select_external_doc(self.pages_path_input))
        pages_path_layout = QtWidgets.QHBoxLayout()
        pages_path_layout.addWidget(self.pages_path_input)
        pages_path_layout.addWidget(pages_path_button)
        pages_layout.addRow("Seiten-Typ:", self.pages_type_combo)
        pages_layout.addRow("Anzahl leerer Seiten:", self.blank_pages_spinbox)
        pages_layout.addRow("Pfad zur DOCX:", pages_path_layout)
        layout.addWidget(pages_group)

        # --- Gruppe für Tabellen-Stile ---
        group = QtWidgets.QGroupBox("Stil-Einstellungen für Katalog-Tabellen")
        group_layout = QtWidgets.QFormLayout(group)
        
        self.style_name_input = QtWidgets.QLineEdit()
        self.style_name_input.setToolTip("Geben Sie den Namen eines in Word verfügbaren Tabellen-Stils ein (z.B. 'Table Grid').")
        group_layout.addRow("Word Tabellen-Stil:", self.style_name_input)
        
        self.header_bold_check = QtWidgets.QCheckBox("Überschrift fett formatieren")
        group_layout.addRow(self.header_bold_check)
        
        self.header_font_color_input = QtWidgets.QLineEdit()
        header_font_color_button = QtWidgets.QPushButton("Farbe auswählen...")
        header_font_color_button.clicked.connect(lambda: self._pick_shading_color(self.header_font_color_input))
        header_font_color_layout = QtWidgets.QHBoxLayout()
        header_font_color_layout.addWidget(self.header_font_color_input)
        header_font_color_layout.addWidget(header_font_color_button)
        group_layout.addRow("Header-Schriftfarbe (HEX):", header_font_color_layout)
        self.header_shading_color_input = QtWidgets.QLineEdit()
        header_color_button = QtWidgets.QPushButton("Farbe auswählen...")
        header_color_button.clicked.connect(lambda: self._pick_shading_color(self.header_shading_color_input))
        header_color_layout = QtWidgets.QHBoxLayout()
        header_color_layout.addWidget(self.header_shading_color_input)
        header_color_layout.addWidget(header_color_button)
        group_layout.addRow("Header-Hintergrundfarbe (HEX):", header_color_layout)

        self.shading_enabled_check = QtWidgets.QCheckBox("Zeilenschattierung (jede zweite Zeile)")
        group_layout.addRow(self.shading_enabled_check)
        
        self.shading_color_input = QtWidgets.QLineEdit()
        self.shading_color_input.setToolTip("Geben Sie die Farbe als Hexadezimalwert ohne '#' an (z.B. 'DAE9F8').")
        
        color_button = QtWidgets.QPushButton("Farbe auswählen...")
        color_button.clicked.connect(lambda: self._pick_shading_color(self.shading_color_input))
        
        color_layout = QtWidgets.QHBoxLayout()
        color_layout.addWidget(self.shading_color_input)
        color_layout.addWidget(color_button)
        group_layout.addRow("Zeilen-Schattierungsfarbe (HEX):", color_layout)
        
        layout.addWidget(group)
        title_group = QtWidgets.QGroupBox("Formatierung der Baugruppen-Seite")
        title_layout = QtWidgets.QFormLayout(title_group)
        self.assembly_title_format_input = QtWidgets.QLineEdit()
        self.assembly_title_format_input.setToolTip("Verfügbare Platzhalter: {benennung}, {teilenummer}")
        title_layout.addRow("Format-String für Überschrift:", self.assembly_title_format_input)
        
        self.space_after_title_check = QtWidgets.QCheckBox("Leerraum zwischen Titel und Grafik einfügen")
        self.table_on_new_page_check = QtWidgets.QCheckBox("Tabelle auf neuer Seite beginnen (nach Grafik)")
        self.toc_on_new_page_check = QtWidgets.QCheckBox("Inhaltsverzeichnis auf eigener Seite beginnen") # NEU
        title_layout.addRow(self.space_after_title_check)
        title_layout.addRow(self.table_on_new_page_check)
        title_layout.addRow(self.toc_on_new_page_check)
        layout.addWidget(title_group)

        layout.addStretch()

    def _setup_mapping_tab(self):
        """Erstellt die UI für den 'Spaltenzuordnung'-Tab."""
        layout = QtWidgets.QVBoxLayout(self.mapping_tab)
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
        self._create_column_comboboxes()

        layout.addWidget(header_group)
        layout.addWidget(self.column_group)
        layout.addStretch()

    def _setup_layout_tab(self):
        """Erstellt die UI für den 'Katalog-Layout'-Tab."""
        layout = QtWidgets.QVBoxLayout(self.layout_tab)
        self.layout_table = QtWidgets.QTableWidget()
        self.layout_table.setColumnCount(3)
        self.layout_table.setHorizontalHeaderLabels(
            ["Spaltenüberschrift", "Datenquelle (interne ID)", "Breite (%)"]
        )
        header_view = self.layout_table.horizontalHeader()
        header_view.setSectionResizeMode(
            0, QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        header_view.setSectionResizeMode(
            1, QtWidgets.QHeaderView.ResizeMode.Stretch
        )
        header_view.setSectionResizeMode(
            2, QtWidgets.QHeaderView.ResizeMode.ResizeToContents
        )
        self._populate_layout_table()

        button_layout = QtWidgets.QHBoxLayout()
        add_button = QtWidgets.QPushButton("Benutzerdefinierte Spalte hinzufügen")
        remove_button = QtWidgets.QPushButton("Ausgewählte Spalte entfernen")
        up_button = QtWidgets.QPushButton("Hoch")
        down_button = QtWidgets.QPushButton("Runter")
        button_layout.addWidget(add_button)
        button_layout.addWidget(remove_button)
        button_layout.addStretch()
        button_layout.addWidget(up_button)
        button_layout.addWidget(down_button)

        layout.addWidget(self.layout_table)
        layout.addLayout(button_layout)

        self.total_width_label = QtWidgets.QLabel()
        self.total_width_label.setAlignment(
            QtCore.Qt.AlignmentFlag.AlignRight
        )
        self.total_width_label.setStyleSheet(
            "font-weight: bold; padding-right: 10px;"
        )
        layout.addWidget(self.total_width_label)

        add_button.clicked.connect(self._add_layout_row)
        remove_button.clicked.connect(self._remove_layout_row)
        up_button.clicked.connect(partial(self._move_layout_row, -1))
        down_button.clicked.connect(partial(self._move_layout_row, 1))

    def _setup_rules_tab(self):
        """Erstellt die UI für den 'Setzregeln'-Tab."""
        main_splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        tab_layout = QtWidgets.QHBoxLayout(self.rules_tab)
        tab_layout.addWidget(main_splitter)

        left_panel = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.addWidget(QtWidgets.QLabel("Generierte Katalog-Felder:"))
        self.target_list = QtWidgets.QListWidget()
        left_layout.addWidget(self.target_list)
        target_button_layout = QtWidgets.QHBoxLayout()
        self.add_target_button = QtWidgets.QPushButton("Neu")
        self.remove_target_button = QtWidgets.QPushButton("Löschen")
        target_button_layout.addWidget(self.add_target_button)
        target_button_layout.addWidget(self.remove_target_button)
        left_layout.addLayout(target_button_layout)
        main_splitter.addWidget(left_panel)

        right_panel = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right_panel)
        right_layout.setContentsMargins(5, 0, 0, 0)
        self.rule_editor_groupbox = QtWidgets.QGroupBox(
            "Regel-Definition für: [Kein Feld ausgewählt]"
        )
        right_layout.addWidget(self.rule_editor_groupbox)
        self._setup_rule_editor_area()
        main_splitter.addWidget(right_panel)
        main_splitter.setSizes([250, 650])

    def _setup_rule_editor_area(self):
        """Erstellt den Inhalt des rechten Panels für den Regeleditor."""
        editor_layout = QtWidgets.QVBoxLayout(self.rule_editor_groupbox)
        rule_type_layout = QtWidgets.QHBoxLayout()
        rule_type_layout.addWidget(QtWidgets.QLabel("Regel-Typ:"))
        self.rule_type_combo = QtWidgets.QComboBox()
        self.rule_type_combo.addItems(self.rule_type_map.keys())
        rule_type_layout.addWidget(self.rule_type_combo)
        editor_layout.addLayout(rule_type_layout)
        self.rule_stack = QtWidgets.QStackedWidget()
        editor_layout.addWidget(self.rule_stack)

        self.rule_stack.addWidget(self._create_list_based_ui("prio"))
        self.rule_stack.addWidget(self._create_combine_ui())
        self.rule_stack.addWidget(self._create_conditional_ui())

    def _create_list_based_ui(self, rule_prefix: str) -> QtWidgets.QWidget:
        """Erstellt UI für 'Priorisierte Liste' und 'Werte kombinieren'."""
        widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(widget)

        if rule_prefix == "prio":
            layout.addWidget(QtWidgets.QLabel("<b>Funktion:</b> Nimmt den Wert des ersten Feldes in der Liste, das nicht leer ist."))
            layout.addWidget(QtWidgets.QLabel("Quell-Felder (in Priorität von oben nach unten):"))
            list_widget = self.prio_list_widget = QtWidgets.QListWidget()
        else:
            layout.addWidget(QtWidgets.QLabel("<b>Funktion:</b> Verbindet die Werte mehrerer Felder mit einem Trennzeichen."))
            layout.addWidget(QtWidgets.QLabel("Zu kombinierende Quell-Felder (in Reihenfolge):"))
            list_widget = self.combine_list_widget = QtWidgets.QListWidget()

        layout.addWidget(list_widget)

        button_layout = QtWidgets.QHBoxLayout()
        add_btn = QtWidgets.QPushButton("Quelle Hinzufügen")
        remove_btn = QtWidgets.QPushButton("Quelle Entfernen")
        up_btn = QtWidgets.QPushButton("▲")
        down_btn = QtWidgets.QPushButton("▼")
        button_layout.addWidget(add_btn)
        button_layout.addWidget(remove_btn)
        button_layout.addStretch()
        button_layout.addWidget(up_btn)
        button_layout.addWidget(down_btn)
        layout.addLayout(button_layout)

        add_btn.clicked.connect(partial(self._add_source_to_list, list_widget))
        remove_btn.clicked.connect(partial(self._remove_source_from_list, list_widget))
        up_btn.clicked.connect(partial(self._move_list_item, list_widget, -1))
        down_btn.clicked.connect(partial(self._move_list_item, list_widget, 1))

        if rule_prefix == "combine":
            separator_layout = QtWidgets.QHBoxLayout()
            separator_layout.addWidget(QtWidgets.QLabel("Trennzeichen:"))
            self.separator_input = QtWidgets.QLineEdit()
            self.separator_input.setToolTip("Für einen Zeilenumbruch '\\n' verwenden.")
            separator_layout.addWidget(self.separator_input)
            layout.addLayout(separator_layout)

        return widget

    def _create_combine_ui(self) -> QtWidgets.QWidget:
        return self._create_list_based_ui("combine")

    def _create_conditional_ui(self) -> QtWidgets.QWidget:
        widget = QtWidgets.QWidget()
        form_layout = QtWidgets.QFormLayout(widget)
        form_layout.setContentsMargins(10, 10, 10, 10)
        form_layout.setSpacing(10)

        form_layout.addRow(QtWidgets.QLabel("<b>WENN-Bedingung:</b>"))
        self.if_source_combo = QtWidgets.QComboBox()
        form_layout.addRow("    Feld:", self.if_source_combo)
        self.if_operator_combo = QtWidgets.QComboBox()
        self.if_operator_combo.addItems(self.operator_map.keys())
        form_layout.addRow("    Operator:", self.if_operator_combo)
        self.if_value_input = QtWidgets.QLineEdit()
        self.if_value_input.setToolTip("Mehrere Werte mit Semikolon (;) trennen.")
        form_layout.addRow("    Vergleichswert:", self.if_value_input)

        form_layout.addRow(QtWidgets.QLabel("<b>DANN-Aktion:</b>"))
        self.then_source_combo = QtWidgets.QComboBox()
        form_layout.addRow("    Nimm Wert aus Feld:", self.then_source_combo)

        form_layout.addRow(QtWidgets.QLabel("<b>SONST-Aktion:</b>"))
        self.else_source_combo = QtWidgets.QComboBox()
        form_layout.addRow("    Nimm Wert aus Feld:", self.else_source_combo)

        self._populate_conditional_combos()
        return widget

    # --------------------------------------------------------------------------
    # --- Signal-Verbindungen und Logik ---
    # --------------------------------------------------------------------------

    def _connect_signals(self):
        """Verbindet alle Signale mit den entsprechenden Slots."""
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        self.target_list.currentItemChanged.connect(self._on_target_selected)
        self.rule_type_combo.currentIndexChanged.connect(self.rule_stack.setCurrentIndex)
        self.add_target_button.clicked.connect(self._add_new_target_field)
        self.remove_target_button.clicked.connect(self._remove_target_field)

    def _on_target_selected(self, current_item, previous_item):
        """Wird aufgerufen, wenn links ein Zielfeld ausgewählt wird."""
        if previous_item:
            self._save_current_rule_state(previous_item.text())

        if not current_item:
            self.rule_editor_groupbox.setTitle(
                "Regel-Definition für: [Kein Feld ausgewählt]"
            )
            self.rule_editor_groupbox.setEnabled(False)
            return

        self.current_target_field = current_item.text()
        self.rule_editor_groupbox.setTitle(
            f"Regel-Definition für: '{self.current_target_field}'"
        )
        self.rule_editor_groupbox.setEnabled(True)
        self._populate_conditional_combos()
        self._load_rule_for_target(self.current_target_field)


    def _load_design_settings(self):
        """Lädt die Design-Einstellungen und befüllt die UI."""
        styles = self.config_manager.config.get("table_styles", {})
        self.style_name_input.setText(styles.get("base_style", "Table Grid"))
        self.header_bold_check.setChecked(styles.get("header_bold", True))
        self.header_font_color_input.setText(styles.get("header_font_color", "FFFFFF"))
        self.header_shading_color_input.setText(styles.get("header_shading_color", "4F81BD"))
        self.shading_enabled_check.setChecked(styles.get("shading_enabled", True))
        self.shading_color_input.setText(styles.get("shading_color", "DAE9F8"))

        formatting = self.config_manager.config.get("formatting_options", {})
        self.assembly_title_format_input.setText(
            formatting.get("assembly_title_format", "{benennung} ({teilenummer})")
        )
        self.space_after_title_check.setChecked(formatting.get("add_space_after_title", True))
        self.table_on_new_page_check.setChecked(formatting.get("table_on_new_page", False))
        self.blank_pages_spinbox.setValue(formatting.get("blank_pages_before_toc", 0))
        self.toc_on_new_page_check.setChecked(formatting.get("toc_on_new_page", True))
        self.cover_type_combo.setCurrentText("Externe DOCX-Datei" if formatting.get("cover_sheet_type") == "external_docx" else "Standard (Titel + Bild)")
        self.cover_path_input.setText(formatting.get("cover_sheet_path", ""))
        self.pages_type_combo.setCurrentText("Externe DOCX-Datei" if formatting.get("blank_pages_type") == "external_docx" else "Leere Seiten")
        self.pages_path_input.setText(formatting.get("blank_pages_path", ""))

    def _load_rule_for_target(self, target_field: str):
        """Lädt die Regel für das ausgewählte Feld und stellt die UI ein."""
        rule = self.current_rules.get(target_field, {})
        rule_type_key = rule.get("type", "prioritized_list")

        ui_text_list = [k for k, v in self.rule_type_map.items() if v == rule_type_key]
        if ui_text_list:
            self.rule_type_combo.setCurrentText(ui_text_list[0])

        if rule_type_key == "prioritized_list":
            self.prio_list_widget.clear()
            self.prio_list_widget.addItems(rule.get("sources", []))
        elif rule_type_key == "combine":
            self.combine_list_widget.clear()
            self.combine_list_widget.addItems(rule.get("sources", []))
            self.separator_input.setText(rule.get("separator", ""))
        elif rule_type_key == "conditional":
            if_clause = rule.get("if", {})
            then_clause = rule.get("then", {})
            else_clause = rule.get("else", {})
            self.if_source_combo.setCurrentText(if_clause.get("source", ""))
            op_ui_text_list = [
                k for k, v in self.operator_map.items() if v == if_clause.get("operator")
            ]
            if op_ui_text_list:
                self.if_operator_combo.setCurrentText(op_ui_text_list[0])
            self.if_value_input.setText(if_clause.get("value", ""))
            self.then_source_combo.setCurrentText(then_clause.get("source", ""))
            self.else_source_combo.setCurrentText(else_clause.get("source", ""))

    def _save_current_rule_state(self, target_field: str):
        """Liest die UI-Werte aus und speichert sie in self.current_rules."""
        if not target_field:
            return
        
        rule_type_key = self.rule_type_map[self.rule_type_combo.currentText()]
        new_rule = {"type": rule_type_key}

        if rule_type_key == "prioritized_list":
            sources = [self.prio_list_widget.item(i).text() for i in range(self.prio_list_widget.count())]
            new_rule["sources"] = sources
        elif rule_type_key == "combine":
            sources = [self.combine_list_widget.item(i).text() for i in range(self.combine_list_widget.count())]
            new_rule["sources"] = sources
            new_rule["separator"] = self.separator_input.text()
        elif rule_type_key == "conditional":
            operator_key = self.operator_map[self.if_operator_combo.currentText()]
            new_rule["if"] = {"source": self.if_source_combo.currentText(), "operator": operator_key, "value": self.if_value_input.text()}
            new_rule["then"] = {"source": self.then_source_combo.currentText()}
            new_rule["else"] = {"source": self.else_source_combo.currentText()}
            
        self.current_rules[target_field] = new_rule

    def accept(self):
        """Sammelt Daten aus allen Tabs und speichert die Konfiguration."""
        if self.current_target_field:
            self._save_current_rule_state(self.current_target_field)

        new_header_mapping = {
            key: widget.text().upper() for key, widget in self.header_widgets.items()
        }
        new_column_mapping = {
            key: widget.currentText().split(" - ")[0] for key, widget in self.column_widgets.items()
        }
        new_output_columns = [
            data for row in range(self.layout_table.rowCount()) if (data := self._get_row_data(row))
        ]

        new_table_styles = {
            "base_style": self.style_name_input.text(),
            "header_bold": self.header_bold_check.isChecked(),
            "header_font_color": self.header_font_color_input.text(),
            "header_shading_color": self.header_shading_color_input.text(),
            "shading_enabled": self.shading_enabled_check.isChecked(),
            "shading_color": self.shading_color_input.text()
        }

        new_formatting_options = {
            "assembly_title_format": self.assembly_title_format_input.text(),
            "add_space_after_title": self.space_after_title_check.isChecked(),
            "table_on_new_page": self.table_on_new_page_check.isChecked(),
            "blank_pages_before_toc": self.blank_pages_spinbox.value(),
            "toc_on_new_page": self.toc_on_new_page_check.isChecked(),
            "cover_sheet_type": "external_docx" if self.cover_type_combo.currentText() == "Externe DOCX-Datei" else "default",
            "cover_sheet_path": self.cover_path_input.text(),
            "blank_pages_type": "external_docx" if self.pages_type_combo.currentText() == "Externe DOCX-Datei" else "blank",
            "blank_pages_path": self.pages_path_input.text()
        }

        new_config = self.config_manager.config
        new_config.update({
            "header_mapping": new_header_mapping,
            "column_mapping": new_column_mapping,
            "output_columns": new_output_columns,

            "generation_rules": self.current_rules,
            "table_styles": new_table_styles,
            "formatting_options": new_formatting_options
        })
        self.config_manager.save_config(new_config)
        super().accept()

    # --------------------------------------------------------------------------
    # --- Hilfsmethoden ---
    # --------------------------------------------------------------------------

    def _select_external_doc(self, target_line_edit: QtWidgets.QLineEdit):
        """Öffnet einen Datei-Dialog zur Auswahl einer DOCX-Datei."""
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "DOCX-Datei auswählen", "", "Word-Dokumente (*.docx)")
        if path:
            target_line_edit.setText(path)

    def _get_all_available_sources(self, exclude_field: str = None) -> list:
        """Sammelt alle verfügbaren Quell-Felder (Rohdaten + andere Regeln)."""
        input_fields = list(self.config_manager.config.get("column_mapping", {}).keys())
        rule_fields = list(self.current_rules.keys())
        all_sources = sorted(list(set(input_fields + rule_fields)))
        if exclude_field and exclude_field in all_sources:
            all_sources.remove(exclude_field)
        return [""] + all_sources

    def _update_total_width(self):
        """Berechnet die Gesamtbreite und aktualisiert das Label."""
        total = sum(
            self.layout_table.cellWidget(row, 2).value()
            for row in range(self.layout_table.rowCount())
            if isinstance(self.layout_table.cellWidget(row, 2), QtWidgets.QSpinBox)
        )
        color = "red" if total != 100 else "black"
        self.total_width_label.setText(f"<b>Gesamtbreite: {total} %</b>")
        self.total_width_label.setStyleSheet(
            f"font-weight: bold; padding-right: 10px; color: {color};"
        )

    def _create_row_widgets(self, row, data, available_ids):
        """Erstellt die Widgets für eine einzelne Zeile der Layout-Tabelle."""
        is_standard = data.get("type") == "standard"
        item_header = QtWidgets.QTableWidgetItem(data.get("header", "Neue Spalte"))
        self.layout_table.setItem(row, 0, item_header)
        item_header.setData(QtCore.Qt.ItemDataRole.UserRole, data)
        source_id = data.get("source_id", "")
        if is_standard:
            label = QtWidgets.QLabel(f"<i>{source_id} (Standard)</i>")
            label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            self.layout_table.setCellWidget(row, 1, label)
        else:
            combo_id = QtWidgets.QComboBox()
            combo_id.addItems(available_ids)
            combo_id.setCurrentText(source_id)
            self.layout_table.setCellWidget(row, 1, combo_id)

        spin_width = QtWidgets.QSpinBox()
        spin_width.setRange(1, 100)
        spin_width.setSuffix(" %")
        spin_width.setValue(int(data.get("width_percent", 10)))
        self.layout_table.setCellWidget(row, 2, spin_width)
        spin_width.valueChanged.connect(self._update_total_width)

    def _add_layout_row(self):
        """Fügt eine neue, leere Zeile zur Layout-Tabelle hinzu."""
        row = self.layout_table.rowCount()
        self.layout_table.insertRow(row)
        available_ids = self._get_all_available_sources()
        new_id = f"custom_{int(time.time() * 1000)}"
        new_col_data = {"id": new_id, "type": "custom", "header": "Neue Spalte", "width_percent": 15, "source_id": ""}
        self._create_row_widgets(row, new_col_data, available_ids)
        self._update_total_width()

    def _remove_layout_row(self):
        """Entfernt die ausgewählte Zeile aus der Layout-Tabelle."""
        row = self.layout_table.currentRow()
        if row >= 0:
            item = self.layout_table.item(row, 0)
            if item:
                stored_data = item.data(QtCore.Qt.ItemDataRole.UserRole)
                if isinstance(stored_data, dict) and stored_data.get("type") == "custom":
                    self.layout_table.removeRow(row)
                    self._update_total_width()
                else:
                    QtWidgets.QMessageBox.warning(self, "Fehler", "Standard-Spalten können nicht entfernt werden.")

    def _move_layout_row(self, direction: int):
        """Bewegt die ausgewählte Zeile nach oben oder unten."""
        row = self.layout_table.currentRow()
        if row == -1: return
        
        new_row = row + direction
        if 0 <= new_row < self.layout_table.rowCount():
            all_sources = self._get_all_available_sources()
            data = self._get_row_data(row)
            self.layout_table.removeRow(row)
            self.layout_table.insertRow(new_row)
            self._create_row_widgets(new_row, data, all_sources)
            self.layout_table.setCurrentCell(new_row, 0)

    def _populate_target_fields_list(self):
        """Füllt die linke Liste der Setzregeln und aktualisiert die Quellen."""
        self.target_list.clear()
        self.target_list.addItems(sorted(self.current_rules.keys()))
        self._populate_conditional_combos()

    def _add_new_target_field(self):
        """Fügt ein neues, leeres generiertes Feld hinzu."""
        text, ok = QtWidgets.QInputDialog.getText(self, "Neues Feld", "Name des neuen generierten Feldes:")
        if ok and text:
            if text in self.current_rules:
                QtWidgets.QMessageBox.warning(self, "Fehler", "Ein Feld mit diesem Namen existiert bereits.")
                return
            self.current_rules[text] = {"type": "prioritized_list", "sources": []}
            self._populate_target_fields_list()

            items = self.target_list.findItems(text, QtCore.Qt.MatchFlag.MatchExactly)
            if items:
                self.target_list.setCurrentItem(items[0])
    
    def _remove_target_field(self):
        """Entfernt das ausgewählte generierte Feld."""
        current_item = self.target_list.currentItem()
        if not current_item: return
        
        reply = QtWidgets.QMessageBox.question(self, "Löschen", f"Möchten Sie das Feld '{current_item.text()}' und seine Regel wirklich löschen?")
        if reply == QtWidgets.QMessageBox.StandardButton.Yes:
            del self.current_rules[current_item.text()]
            self._populate_target_fields_list()

    def _add_source_to_list(self, list_widget: QtWidgets.QListWidget):
        """Fügt eine ausgewählte Quelle zu einer der Listen-Widgets hinzu."""
        available_sources = self._get_all_available_sources(self.current_target_field)
        source, ok = QtWidgets.QInputDialog.getItem(self, "Quelle auswählen", "Wählen Sie ein Quell-Feld aus:", available_sources, 0, False)
        if ok and source:
            list_widget.addItem(source)

    def _remove_source_from_list(self, list_widget: QtWidgets.QListWidget):
        """Entfernt die ausgewählte Quelle aus einem Listen-Widget."""
        current_item = list_widget.currentItem()
        if current_item:
            list_widget.takeItem(list_widget.row(current_item))
            
    def _move_list_item(self, list_widget: QtWidgets.QListWidget, direction: int):
        """Bewegt ein Item in einem Listen-Widget nach oben oder unten."""
        current_row = list_widget.currentRow()
        if current_row == -1: return
        new_row = current_row + direction
        if 0 <= new_row < list_widget.count():
            item = list_widget.takeItem(current_row)
            list_widget.insertItem(new_row, item)
            list_widget.setCurrentRow(new_row)

    def _populate_conditional_combos(self):
        """Füllt die Dropdowns für die bedingte Zuweisung neu."""
        available_sources = self._get_all_available_sources(self.current_target_field)
        self.if_source_combo.clear()
        self.if_source_combo.addItems(available_sources)
        self.then_source_combo.clear()
        self.then_source_combo.addItems(available_sources)
        self.else_source_combo.clear()
        self.else_source_combo.addItems(available_sources)
            

    def _create_column_comboboxes(self):
        """Erstellt oder aktualisiert die Dropdowns für die Spaltenzuordnung."""

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
            full_text_to_set = next((opt for opt in self.available_columns if opt.startswith(current_value + " -")), None)
            if full_text_to_set:
                combo_box.setCurrentText(full_text_to_set)
            else:
                combo_box.addItem(current_value)
                combo_box.setCurrentText(current_value)
            self.column_widgets[key] = combo_box
            self.column_layout.addWidget(label, row, 0)
            self.column_layout.addWidget(combo_box, row, 1)
            
    def _populate_layout_table(self):

        """Füllt die Layout-Tabelle mit Daten aus der Konfiguration."""

        output_config = self.config_manager.config.get("output_columns", [])
        available_ids = self._get_all_available_sources()
        self.layout_table.setRowCount(len(output_config))
        for row_idx, col_data in enumerate(output_config):
            self._create_row_widgets(row_idx, col_data, available_ids)

    def _get_row_data(self, row):

        """Liest die Daten aus einer Zeile der Layout-Tabelle."""
  
        header_item = self.layout_table.item(row, 0)
        if not header_item: return None
        stored_data = header_item.data(QtCore.Qt.ItemDataRole.UserRole)
        if not isinstance(stored_data, dict):
            stored_data = {"id": None, "type": "custom"}
        col_type = stored_data.get("type")
        col_id = stored_data.get("id")
        source_widget = self.layout_table.cellWidget(row, 1)
        source_id = ""
        if isinstance(source_widget, QtWidgets.QComboBox):
            source_id = source_widget.currentText()
        elif isinstance(source_widget, QtWidgets.QLabel):
            match = re.search(r'<i>(.*?) \(Standard\)</i>', source_widget.text())
            source_id = match.group(1) if match else ''
        width_widget = self.layout_table.cellWidget(row, 2)
        return {"id": col_id, "header": header_item.text(), "width_percent": width_widget.value(), "source_id": source_id, "type": col_type}
        
    def _pick_shading_color(self, target_line_edit):
        """Öffnet einen Farbdialog und setzt den Hex-Wert im Ziel-LineEdit."""
        current_color = target_line_edit.text()
        dialog = QtWidgets.QColorDialog(self)
        if QtGui.QColor.isValidColor(f"#{current_color}"):
            dialog.setCurrentColor(QtGui.QColor(f"#{current_color}"))
        
        if dialog.exec():
            color = dialog.selectedColor()
            target_line_edit.setText(color.name().replace("#", "").upper())