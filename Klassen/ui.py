from PySide6 import QtWidgets, QtCore, QtGui
import os
import sys
import shutil
import json
from functools import partial
from Klassen.config import ConfigManager
from Klassen.generator import DocxGenerator
from Klassen.stueckliste import BomProcessor
from Klassen.editor_ui import ConfigEditorWindow

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, project_path: str):
        super().__init__()
        self.project_path = project_path
        self.config = ConfigManager(project_path)
        self.all_boms = {}
        self.item_lookup = {}

        self.setWindowTitle(f"Ersatzteilkatalog-Generator - Projekt: {os.path.basename(self.project_path)}")
        self.resize(1200, 800)
        self._setup_ui()
        self._connect_signals()
        self._initialize_project()

    def _setup_ui(self):
        """Erstellt alle UI-Elemente und ordnet sie im Layout an."""
        self.assembly_selector = QtWidgets.QComboBox()
        self.author_input = QtWidgets.QLineEdit()
        self.author_input.setPlaceholderText("Ihr Name")
        self.cover_graphic_button = QtWidgets.QPushButton("Deckblatt-Grafik...")
        self.tree_widget = QtWidgets.QTreeWidget()
        self.tree_widget.setColumnCount(6)
        self.tree_widget.setHeaderLabels(["Benennung", "Pos.", "Menge", "Bestellnummer", "Information", "Grafik"])
        col_widths = [350, 50, 80, 180, 250, 100]
        for i, width in enumerate(col_widths): 
            self.tree_widget.setColumnWidth(i, width)

        self.generate_button = QtWidgets.QPushButton("Katalog generieren")
        self.update_fields_checkbox = QtWidgets.QCheckBox("Felder auto. aktualisieren (benötigt MS Word)")
        if sys.platform == "win32":
            self.update_fields_checkbox.setChecked(True)
        else: 
            self.update_fields_checkbox.setVisible(False)
            self.update_fields_checkbox.setChecked(False)
        self.save_button = QtWidgets.QPushButton("Auswahl speichern")
        self.load_button = QtWidgets.QPushButton("Auswahl laden")
        self.info_button = QtWidgets.QPushButton("Info / Copyright")
        self.config_button = QtWidgets.QPushButton("Spalten-Editor")

        # --- Layout ---
        top_layout = QtWidgets.QHBoxLayout()
        top_layout.addWidget(QtWidgets.QLabel("Hauptbaugruppe:"))
        top_layout.addWidget(self.assembly_selector, 2)
        top_layout.addWidget(QtWidgets.QLabel("Ersteller:")); top_layout.addWidget(self.author_input, 1)
        top_layout.addWidget(self.cover_graphic_button)
        
        bottom_layout = QtWidgets.QHBoxLayout()
        bottom_layout.addWidget(self.info_button)
        bottom_layout.addWidget(self.config_button)
        bottom_layout.addWidget(self.load_button)
        bottom_layout.addWidget(self.save_button)
        bottom_layout.addStretch()
        bottom_layout.addWidget(self.update_fields_checkbox)
        bottom_layout.addWidget(self.generate_button)
        
        main_layout = QtWidgets.QVBoxLayout()
        main_layout.addLayout(top_layout)
        main_layout.addWidget(self.tree_widget)
        main_layout.addLayout(bottom_layout)
        central_widget = QtWidgets.QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

    def _connect_signals(self):
        """Verbindet die Signale der UI-Elemente mit ihren Funktionen."""
        self.assembly_selector.currentTextChanged.connect(self._on_assembly_selected)
        self.generate_button.clicked.connect(self._on_generate_button_clicked)
        self.save_button.clicked.connect(self._on_save_selection_clicked)
        self.load_button.clicked.connect(self._on_load_selection_clicked)
        self.cover_graphic_button.clicked.connect(self._on_assign_cover_graphic_clicked)
        self.tree_widget.itemChanged.connect(self._handle_item_changed)
        self.info_button.clicked.connect(self._show_info_dialog) 
        self.config_button.clicked.connect(self._open_config_editor)

    def _open_config_editor(self):
        """Öffnet den Konfigurations-Editor-Dialog."""
        editor_dialog = ConfigEditorWindow(self.config, self)
        # .exec() öffnet den Dialog modal (blockiert das Hauptfenster)
        if editor_dialog.exec(): # Entspricht QDialog.DialogCode.Accepted
            print("INFO: Konfiguration wurde gespeichert. Lade Projekt neu...")
            # Lade die Projektdaten neu, um die Änderungen zu übernehmen.
            self._load_project_data()
            QtWidgets.QMessageBox.information(self, "Konfiguration gespeichert", 
                "Die Spaltenzuordnung wurde aktualisiert. Das Projekt wurde neu geladen.")

    # NEU: Methode zum Anzeigen des Info-Dialogs
    def _show_info_dialog(self):
        """Zeigt ein Fenster mit Copyright-Informationen an."""
        msg_box = QtWidgets.QMessageBox(self)
        msg_box.setWindowTitle("Information und Copyright")
        msg_box.setIcon(QtWidgets.QMessageBox.Icon.Information)
        msg_box.setText(
            "<b>Ersatzteilkatalog-Generator</b><br><br>"
            "Alle Rechte für diese Software liegen bei:<br>"
            "<b>Marcus Kohtz (Signz-vision.de)</b>"
        )
        msg_box.setInformativeText(
            "Eine Weitergabe, Vervielfältigung oder Nutzung dieser Anwendung, "
            "ganz oder in Teilen, ist ohne die ausdrückliche schriftliche "
            "Genehmigung des Urhebers nicht gestattet."
        )
        msg_box.setStandardButtons(QtWidgets.QMessageBox.StandardButton.Ok)
        msg_box.exec()

    # Der Rest der Klasse bleibt unverändert
    def _initialize_project(self):
        boms_path = os.path.join(self.project_path, "stücklisten")
        if not os.path.isdir(boms_path):
            reply = QtWidgets.QMessageBox.question(self, 'Neues Projekt', 
                "Der ausgewählte Ordner scheint kein Projekt zu sein. Möchten Sie hier ein neues Projekt erstellen?",
                QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No, 
                QtWidgets.QMessageBox.StandardButton.Yes)
            if reply == QtWidgets.QMessageBox.StandardButton.Yes: self._setup_new_project(boms_path)
            else: self.close(); return
        else:
            self._load_project_data()
            self._auto_load_save_file()
    def _setup_new_project(self, boms_path):
        try:
            os.makedirs(boms_path, exist_ok=True)
            os.makedirs(os.path.join(self.project_path, "Grafik"), exist_ok=True)
            master_template_folder = "Vorlagen"
            master_template_path = os.path.join(master_template_folder, "DOK-Vorlage.docm")
            if not os.path.exists(master_template_path): master_template_path = os.path.join(master_template_folder, "DOK-Vorlage.docx")
            if not os.path.exists(master_template_path):
                QtWidgets.QMessageBox.critical(self, "Fehler", "Keine Master-Vorlage im Ordner 'Vorlagen' gefunden."); return
            shutil.copy(master_template_path, self.project_path)
            self._prompt_for_boms(boms_path)
            self._load_project_data()
            self._on_save_selection_clicked(is_initial_save=True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Fehler beim Erstellen des Projekts", str(e))
    def _prompt_for_boms(self, boms_path):
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "Wählen Sie Stücklisten für das neue Projekt", "", "Excel-Dateien (*.xlsm *.xlsx)")
        if not files: QtWidgets.QMessageBox.warning(self, "Abbruch", "Keine Stücklisten ausgewählt. Das Projekt ist leer."); return
        for file_path in files: shutil.copy(file_path, boms_path)
    def _auto_load_save_file(self):
        for filename in os.listdir(self.project_path):
            if filename.lower().startswith("projekt_") and filename.lower().endswith(".json"):
                self._load_selection_from_file(os.path.join(self.project_path, filename)); return
    def _load_selection_from_file(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f: load_data = json.load(f)
            self.author_input.setText(load_data.get('author', ''))
            main_bom_znr = load_data.get('main_bom_znr')
            if main_bom_znr in self.all_boms:
                self.assembly_selector.setCurrentText(main_bom_znr)
                QtCore.QTimer.singleShot(100, lambda: self._apply_loaded_selection(load_data.get('unchecked_item_ids', [])))
            else:
                QtWidgets.QMessageBox.warning(self, "Warnung", f"Die in '{os.path.basename(file_path)}' gespeicherte Hauptbaugruppe '{main_bom_znr}' wurde nicht gefunden.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Fehler beim Laden", f"Die Speicherdatei konnte nicht geladen werden:\n{e}")
    def _load_project_data(self):
        """Liest die Stücklisten und übergibt die Konfiguration."""
        boms_path = os.path.join(self.project_path, "stücklisten")
        if not os.path.isdir(boms_path): 
            return
        processor = BomProcessor(folder_path=boms_path, config_manager=self.config)
        self.all_boms = processor.run()
        self.load_data(self.all_boms)
    def _on_load_selection_clicked(self):
        load_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Auswahl laden", self.project_path, "JSON-Dateien (*.json)")
        if load_path: self._load_selection_from_file(load_path)
    def _on_save_selection_clicked(self, is_initial_save=False):
        if not self.assembly_selector.currentText(): return
        unchecked_items = []
        root = self.tree_widget.invisibleRootItem()
        self._collect_unchecked_items(root, unchecked_items)
        save_data = {'main_bom_znr': self.assembly_selector.currentText(), 'author': self.author_input.text(), 'unchecked_item_ids': unchecked_items}
        if is_initial_save:
            save_path = os.path.join(self.project_path, f"projekt_{os.path.basename(self.project_path)}.json")
        else:
            save_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Auswahl speichern", self.project_path, "JSON-Dateien (*.json)")
        if save_path:
            with open(save_path, 'w', encoding='utf-8') as f: json.dump(save_data, f, indent=4, ensure_ascii=False)
            if not is_initial_save:
                QtWidgets.QMessageBox.information(self, "Erfolg", "Die aktuelle Auswahl wurde gespeichert.")
    def _collect_unchecked_items(self, parent_item, collection: list):
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            if child.checkState(0) == QtCore.Qt.CheckState.Unchecked:
                item_data = child.data(0, QtCore.Qt.ItemDataRole.UserRole)
                if item_data and item_data.get('unique_id'): collection.append(item_data.get('unique_id'))
            if child.childCount() > 0: self._collect_unchecked_items(child, collection)
    def _apply_loaded_selection(self, unchecked_id_list: list):
        self.tree_widget.blockSignals(True)
        for unique_id in unchecked_id_list:
            if unique_id in self.item_lookup: self.item_lookup[unique_id].setCheckState(0, QtCore.Qt.CheckState.Unchecked)
        self.tree_widget.blockSignals(False)
    def _on_assembly_selected(self, bom_znr: str):
        if bom_znr in self.all_boms:
            self._populate_tree(self.all_boms[bom_znr])
    def _populate_tree(self, main_assembly):
        self.tree_widget.blockSignals(True)
        self.tree_widget.clear()
        self.item_lookup.clear()
        root_item = QtWidgets.QTreeWidgetItem(self.tree_widget)
        root_item.setText(0, f"{main_assembly.zeichnungsnummer} ({main_assembly.titel})")
        root_item.setFont(0, self._get_bold_font())
        root_item.setFlags(root_item.flags() | QtCore.Qt.ItemFlag.ItemIsUserCheckable)
        root_item.setCheckState(0, QtCore.Qt.CheckState.Checked)
        unique_id = str(main_assembly.zeichnungsnummer)
        root_item_data = {'Benennung': main_assembly.titel, 'Teilenummer': main_assembly.zeichnungsnummer, 'is_assembly': True, 'unique_id': unique_id}; 
        root_item.setData(0, QtCore.Qt.ItemDataRole.UserRole, root_item_data)
        self.item_lookup[unique_id] = root_item
        assign_button = QtWidgets.QPushButton("Zuordnen...")
        assign_button.clicked.connect(partial(self._on_assign_graphic_clicked, root_item))
        self.tree_widget.setItemWidget(root_item, 5, assign_button)
        self._add_children_recursively(root_item, main_assembly)
        self.tree_widget.expandAll()
        self.tree_widget.blockSignals(False)
    def _add_children_recursively(self, parent_item, bom_obj):
        for position in bom_obj.positionen:
            child_item = QtWidgets.QTreeWidgetItem(parent_item)
            child_item.setText(0, str(position.get('Benennung', '')))
            child_item.setText(1, f"{position.get('POS', ''):g}")
            child_item.setText(2, str(position.get('Menge', '')))
            child_item.setText(3, str(position.get('Bestellnummer_Kunde')))
            child_item.setText(4, str(position.get('Information', '')))
            child_item.setFlags(child_item.flags() | QtCore.Qt.ItemFlag.ItemIsUserCheckable)
            child_item.setCheckState(0, QtCore.Qt.CheckState.Checked)
            parent_znr = str(bom_obj.zeichnungsnummer)
            pos_nr = str(position.get('POS', ''))
            unique_id = f"{parent_znr}_{pos_nr}"
            position_data = position.copy()
            position_data['is_assembly'] = 'sub_assembly' in position
            position_data['unique_id'] = unique_id
            child_item.setData(0, QtCore.Qt.ItemDataRole.UserRole, position_data)
            self.item_lookup[unique_id] = child_item
            if position_data['is_assembly']:
                child_item.setFont(0, self._get_bold_font())
                assign_button = QtWidgets.QPushButton("Zuordnen...")
                assign_button.clicked.connect(partial(self._on_assign_graphic_clicked, child_item))
                self.tree_widget.setItemWidget(child_item, 5, assign_button)
                self._add_children_recursively(child_item, position['sub_assembly'])
    def _on_assign_cover_graphic_clicked(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Deckblatt-Grafik auswählen", self.project_path, "Bilder (*.png *.jpg *.jpeg)")
        if file_path:
            grafik_folder = os.path.join(self.project_path, "Grafik")
            os.makedirs(grafik_folder, exist_ok=True)
            _, file_extension = os.path.splitext(file_path)
            destination_path = os.path.join(grafik_folder, f"EL{file_extension}")
            try:
                shutil.copy(file_path, destination_path)
                QtWidgets.QMessageBox.information(self, "Erfolg", "Deckblatt-Grafik wurde erfolgreich gespeichert.")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Fehler beim Kopieren", f"Die Grafik konnte nicht gespeichert werden:\n{e}")
    def _on_assign_graphic_clicked(self, item):
        item_data = item.data(0, QtCore.Qt.ItemDataRole.UserRole)
        if not item_data:
            return
        zeichnung_nr = item_data.get('Teilenummer')
        if not zeichnung_nr:
            QtWidgets.QMessageBox.warning(self, "Fehler", "Für diese Position wurde keine Zeichnungsnummer gefunden.")
            return
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Grafik auswählen", self.project_path, "Bilder (*.png *.jpg *.jpeg)")
        if file_path:
            grafik_folder = os.path.join(self.project_path, "Grafik")
            os.makedirs(grafik_folder, exist_ok=True)
            _, file_extension = os.path.splitext(file_path)
            destination_path = os.path.join(grafik_folder, f"{zeichnung_nr}{file_extension}")
            try:
                shutil.copy(file_path, destination_path)
                QtWidgets.QMessageBox.information(self, "Erfolg", f"Grafik wurde erfolgreich für {zeichnung_nr} gespeichert.")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Fehler beim Kopieren", f"Die Grafik konnte nicht gespeichert werden:\n{e}")
    def _on_generate_button_clicked(self):
        root = self.tree_widget.invisibleRootItem()
        if root.childCount() == 0: return
        root_item = root.child(0)
        if not root_item or root_item.checkState(0) == QtCore.Qt.CheckState.Unchecked:
            return
        hierarchical_data = self._collect_hierarchical_data(root_item)
        if not hierarchical_data or (not hierarchical_data.get('children') and not hierarchical_data.get('is_assembly')):
            return
        main_bom_znr = self.assembly_selector.currentText()
        main_bom_obj = self.all_boms.get(main_bom_znr)
        if not main_bom_obj:
            return
        author_name = self.author_input.text()
        auto_update = self.update_fields_checkbox.isChecked()
        template_path = os.path.join(self.project_path, "DOK-Vorlage.docm")
        if not os.path.exists(template_path):
             template_path = os.path.join(self.project_path, "DOK-Vorlage.docx")
             if not os.path.exists(template_path):
                 QtWidgets.QMessageBox.critical(self, "Fehler", "Keine DOK-Vorlage (.docm oder .docx) im Projektordner gefunden."); return
        suggested_filename = f"Ersatzteilkatalog_{main_bom_znr}.docx"
        output_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Katalog speichern", os.path.join(self.project_path, suggested_filename), "Word-Dokumente (*.docx)")
        if not output_path:
            return
        generator = DocxGenerator(data=hierarchical_data, main_bom=main_bom_obj, author_name=author_name, template_path=template_path, output_path=output_path, auto_update_fields=auto_update, project_path=self.project_path, config_manager=self.config)
        success = generator.run()
        if success:
            msg_box = QtWidgets.QMessageBox()
            msg_box.setIcon(QtWidgets.QMessageBox.Icon.Information)
            msg_box.setText("Erfolg!")
            info_text = f"Der Katalog wurde erfolgreich gespeichert."
            if not auto_update and sys.platform == "win32": info_text += "\n\nWICHTIG: Um den Schriftkopf und das Inhaltsverzeichnis zu aktualisieren, öffnen Sie das Dokument, drücken Sie Strg+A und dann F9."
            msg_box.setInformativeText(info_text)
            msg_box.setStandardButtons(QtWidgets.QMessageBox.StandardButton.Ok | QtWidgets.QMessageBox.StandardButton.Open)
            button_open = msg_box.button(QtWidgets.QMessageBox.StandardButton.Open); button_open.setText("Datei öffnen"); msg_box.exec()
            if msg_box.clickedButton() == button_open: 
                try:
                    os.startfile(output_path)
                except AttributeError:
                    print(f"Datei kann nicht automatisch geöffnet werden. Pfad: {output_path}")
        else: QtWidgets.QMessageBox.critical(self, "Fehler", "Beim Erstellen des Katalogs ist ein Fehler aufgetreten.")
    def load_data(self, all_boms: dict):
        self.all_boms = all_boms
        self.assembly_selector.blockSignals(True)
        self.assembly_selector.clear()
        if not self.all_boms: 
            self.assembly_selector.addItem("Keine Stücklisten gefunden")
            self.assembly_selector.setEnabled(False)
            self.tree_widget.clear()
        else:
            self.assembly_selector.addItems(sorted(self.all_boms.keys()))
            self.assembly_selector.setEnabled(True)
        self.assembly_selector.blockSignals(False)
        if self.assembly_selector.count() > 0:
            self._on_assembly_selected(self.assembly_selector.currentText())
    def _collect_hierarchical_data(self, parent_item):
        if parent_item is None or parent_item.checkState(0) == QtCore.Qt.CheckState.Unchecked:
            return None
        parent_data = parent_item.data(0, QtCore.Qt.ItemDataRole.UserRole)
        if not parent_data:
            return None
        parent_data.pop('sub_assembly', None)
        parent_data['children'] = []
        for i in range(parent_item.childCount()):
            child_item = parent_item.child(i)
            child_data = self._collect_hierarchical_data(child_item)
            if child_data:
                parent_data['children'].append(child_data)
        return parent_data
    def _handle_item_changed(self, item, column):
        if column == 0:
            self.tree_widget.blockSignals(True)
            self._set_children_checkstate(item, item.checkState(0))
            self.tree_widget.blockSignals(False)
    def _set_children_checkstate(self, parent_item, state):
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            child.setCheckState(0, state)
            if child.childCount() > 0:
                self._set_children_checkstate(child, state)
    def _get_bold_font(self):
        font = QtGui.QFont()
        font.setBold(True)
        return font