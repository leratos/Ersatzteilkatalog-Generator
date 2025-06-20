# -*- coding: utf-8 -*-
"""
Dieses Modul definiert das UI-Fenster für den Setzregel-Editor.

Es ermöglicht dem Benutzer die Erstellung und Bearbeitung von Logik-Regeln
für die dynamische Befüllung von Katalogfeldern.
"""

from PySide6 import QtWidgets, QtCore


class RuleEditorWindow(QtWidgets.QDialog):
    """
    Ein Dialogfenster zur Bearbeitung der 'generation_rules' in der
    Projekt-Konfiguration.
    """
    def __init__(self, config_manager, parent=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.current_rules = self.config_manager.config.get("generation_rules", {})
        self.all_source_fields = self._get_all_source_fields()

        self.setWindowTitle("Setzregel-Editor")
        self.setMinimumSize(900, 700)
        self._setup_ui()
        self._connect_signals()
        self._populate_target_fields_list()

    def _setup_ui(self):
        """Erstellt die Haupt-UI-Struktur des Editors."""
        # Hauptlayout (Horizontal Splitter)
        main_splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        self.setLayout(QtWidgets.QVBoxLayout())
        self.layout().addWidget(main_splitter)

        # --- Linkes Panel: Liste der Ziel-Felder ---
        left_panel = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)

        self.target_list = QtWidgets.QListWidget()
        self.target_list.setToolTip("Wählen Sie ein Feld aus, um dessen Regel zu bearbeiten.")
        left_layout.addWidget(QtWidgets.QLabel("Zu befüllende Katalog-Felder:"))
        left_layout.addWidget(self.target_list)

        target_button_layout = QtWidgets.QHBoxLayout()
        self.add_target_button = QtWidgets.QPushButton("Neu")
        self.remove_target_button = QtWidgets.QPushButton("Löschen")
        target_button_layout.addWidget(self.add_target_button)
        target_button_layout.addWidget(self.remove_target_button)
        left_layout.addLayout(target_button_layout)

        main_splitter.addWidget(left_panel)

        # --- Rechtes Panel: Der eigentliche Regel-Editor ---
        right_panel = QtWidgets.QWidget()
        self.right_layout = QtWidgets.QVBoxLayout(right_panel)
        self.right_layout.setContentsMargins(5, 0, 0, 0)
        self.rule_editor_groupbox = QtWidgets.QGroupBox("Regel-Definition für: [Kein Feld ausgewählt]")
        self.right_layout.addWidget(self.rule_editor_groupbox)

        self._setup_rule_editor_area()
        main_splitter.addWidget(right_panel)
        main_splitter.setSizes([250, 650])

        # --- Untere Buttons: Speichern / Abbrechen ---
        self.button_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Save |
            QtWidgets.QDialogButtonBox.StandardButton.Cancel
        )
        self.layout().addWidget(self.button_box)

    def _setup_rule_editor_area(self):
        """Erstellt den Inhalt des rechten Panels."""
        editor_layout = QtWidgets.QVBoxLayout(self.rule_editor_groupbox)

        # Dropdown zur Auswahl des Regel-Typs
        rule_type_layout = QtWidgets.QHBoxLayout()
        rule_type_layout.addWidget(QtWidgets.QLabel("Regel-Typ:"))
        self.rule_type_combo = QtWidgets.QComboBox()
        self.rule_type_combo.addItems([
            "Priorisierte Liste",
            "Werte kombinieren",
            "Bedingte Zuweisung"
        ])
        rule_type_layout.addWidget(self.rule_type_combo)
        editor_layout.addLayout(rule_type_layout)

        # QStackedWidget zur Anzeige der passenden UI für den Regel-Typ
        self.rule_stack = QtWidgets.QStackedWidget()
        editor_layout.addWidget(self.rule_stack)

        # Erstelle die UI für jeden Regel-Typ
        self.rule_stack.addWidget(self._create_prioritized_list_ui())
        self.rule_stack.addWidget(self._create_combine_ui())
        self.rule_stack.addWidget(self._create_conditional_ui())

    # --- UI-Erstellungs-Methoden für die einzelnen Regel-Typen ---

    def _create_prioritized_list_ui(self) -> QtWidgets.QWidget:
        """Erstellt die UI für 'Priorisierte Liste'."""
        widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(widget)
        layout.addWidget(QtWidgets.QLabel("<b>Funktion:</b> Nimmt den Wert des ersten Feldes in der Liste, das nicht leer ist."))
        layout.addWidget(QtWidgets.QLabel("Quell-Felder (in Priorität von oben nach unten):"))

        self.prio_list_widget = QtWidgets.QListWidget()
        self.prio_list_widget.setToolTip("Fügen Sie hier die Felder hinzu, die der Reihe nach geprüft werden sollen.")
        layout.addWidget(self.prio_list_widget)

        # TODO: Buttons zum Hinzufügen, Entfernen und Verschieben von Quellen hinzufügen.
        return widget

    def _create_combine_ui(self) -> QtWidgets.QWidget:
        """Erstellt die UI für 'Werte kombinieren'."""
        widget = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(widget)
        layout.addWidget(QtWidgets.QLabel("<b>Funktion:</b> Verbindet die Werte mehrerer Felder mit einem Trennzeichen."))
        layout.addWidget(QtWidgets.QLabel("Zu kombinierende Quell-Felder (in Reihenfolge):"))

        self.combine_list_widget = QtWidgets.QListWidget()
        self.combine_list_widget.setToolTip("Fügen Sie hier die Felder hinzu, die kombiniert werden sollen.")
        layout.addWidget(self.combine_list_widget)

        separator_layout = QtWidgets.QHBoxLayout()
        separator_layout.addWidget(QtWidgets.QLabel("Trennzeichen:"))
        self.separator_input = QtWidgets.QLineEdit()
        self.separator_input.setToolTip("Geben Sie hier das Trennzeichen ein. Für einen Zeilenumbruch '\\n' verwenden.")
        separator_layout.addWidget(self.separator_input)
        layout.addLayout(separator_layout)

        # TODO: Buttons zum Hinzufügen, Entfernen und Verschieben von Quellen hinzufügen.
        return widget

    def _create_conditional_ui(self) -> QtWidgets.QWidget:
        """Erstellt die UI für 'Bedingte Zuweisung'."""
        widget = QtWidgets.QWidget()
        form_layout = QtWidgets.QFormLayout(widget)
        form_layout.setContentsMargins(10, 10, 10, 10)
        form_layout.setSpacing(10)

        # --- WENN-Klausel ---
        form_layout.addRow(QtWidgets.QLabel("<b>WENN-Bedingung:</b>"))
        self.if_source_combo = QtWidgets.QComboBox()
        self.if_source_combo.addItems(self.all_source_fields)
        form_layout.addRow("    Feld:", self.if_source_combo)

        self.if_operator_combo = QtWidgets.QComboBox()
        self.if_operator_combo.addItems(["ist gleich", "ist nicht gleich", "ist leer", "ist nicht leer", "enthält"])
        form_layout.addRow("    Operator:", self.if_operator_combo)

        self.if_value_input = QtWidgets.QLineEdit()
        self.if_value_input.setToolTip("Bei 'ist gleich' und 'ist nicht gleich' können mehrere Werte mit Semikolon (;) getrennt werden.")
        form_layout.addRow("    Vergleichswert:", self.if_value_input)

        # --- DANN-Klausel ---
        form_layout.addRow(QtWidgets.QLabel("<b>DANN-Aktion:</b>"))
        self.then_source_combo = QtWidgets.QComboBox()
        self.then_source_combo.addItems(self.all_source_fields)
        form_layout.addRow("    Nimm Wert aus Feld:", self.then_source_combo)

        # --- SONST-Klausel ---
        form_layout.addRow(QtWidgets.QLabel("<b>SONST-Aktion:</b>"))
        self.else_source_combo = QtWidgets.QComboBox()
        self.else_source_combo.addItems(self.all_source_fields)
        form_layout.addRow("    Nimm Wert aus Feld:", self.else_source_combo)

        return widget

    # --- Logik-Methoden (werden im nächsten Schritt implementiert) ---

    def _get_all_source_fields(self) -> list:
        """Sammelt alle verfügbaren Felder aus der Konfiguration."""
        # Diese Methode wird später durch die aus dem ConfigManager ersetzt,
        # aber für die UI-Erstellung ist sie hier ausreichend.
        input_fields = list(self.config_manager.config.get("column_mapping", {}).keys())
        return [""] + sorted(input_fields)

    def _populate_target_fields_list(self):
        """Füllt die linke Liste mit den zu bearbeitenden Feldern."""
        self.target_list.clear()
        self.target_list.addItems(sorted(self.current_rules.keys()))

    def _connect_signals(self):
        """Verbindet alle Signale mit den entsprechenden Slots."""
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.target_list.currentItemChanged.connect(self._on_target_selected)
        self.rule_type_combo.currentIndexChanged.connect(self.rule_stack.setCurrentIndex)

    def _on_target_selected(self, current_item, previous_item):
        """Wird aufgerufen, wenn links ein Zielfeld ausgewählt wird."""
        if not current_item:
            self.rule_editor_groupbox.setTitle("Regel-Definition für: [Kein Feld ausgewählt]")
            return

        target_field = current_item.text()
        self.rule_editor_groupbox.setTitle(f"Regel-Definition für: '{target_field}'")
        self._load_rule_for_target(target_field)

    def _load_rule_for_target(self, target_field: str):
        """Lädt die Regel für das ausgewählte Feld und stellt die UI ein."""
        # TODO: Implementierung im nächsten Schritt
        print(f"Lade Regel für '{target_field}'...")
        rule = self.current_rules.get(target_field, {})
        rule_type = rule.get("type")
        # Logik zum Einstellen der Widgets...

    def accept(self):
        """Wird aufgerufen, wenn der Benutzer auf 'Speichern' klickt."""
        # TODO: Implementierung im nächsten Schritt
        print("Speichere alle Regeln...")
        # Logik zum Auslesen aller UI-Werte und Speichern in self.current_rules
        self.config_manager.config["generation_rules"] = self.current_rules
        self.config_manager.save_config(self.config_manager.config)
        super().accept()

