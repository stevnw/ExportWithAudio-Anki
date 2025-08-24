"""

Anki Add-On for exporting notes with audio files to TSV or CSV format.
Spreadsheet GUI with tick boxes to exclude the content you dont want.
Allows for column reordering - which will persist into your exported format. 
Audio is exported and given the file path to the named folder.
Both the audio folder and the tsv/csv are saved into a location of your choosing.

	created by stevnw (https://github.com/stevnw)
	
	Note: 
		- Some of the code is a tad jank and maybe not the best way to do it
		  but it gets the job done. I have only tested this on Debian 12. 
		- This code is provided as is with no guarentee of patches for
		  future version of Anki.

"""

from aqt import mw
from aqt.qt import *
from aqt.utils import showInfo, getSaveFile
from anki.hooks import addHook
import os
import shutil
import csv

def setupMenu():
    action = QAction("Export Notes with Audio", mw)
    action.triggered.connect(onExport)
    mw.form.menuTools.addAction(action)

def onExport():
    deck_id = mw.col.decks.current()['id']
    card_ids = mw.col.decks.cids(deck_id, children=True)
    
    if not card_ids:
        showInfo("No cards found in the current deck.")
        return
    
    selected_nids = list(set(mw.col.get_card(cid).nid for cid in card_ids))
    
    dialog = ExportDialog(mw, selected_nids)
    dialog.exec()

class ExportDialog(QDialog):
    def __init__(self, parent, nids):
        super().__init__(parent)
        self.nids = nids
        self.col = parent.col
        self.field_order = []
        self.setupUI()
        self.loadNoteInfo()
        
    def setupUI(self):
        self.setWindowTitle("Export Notes with Audio")
        self.setMinimumSize(1000, 700)
        
        layout = QHBoxLayout()
        
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 10, 0)
        
        left_label = QLabel("Select columns to export:")
        left_layout.addWidget(left_label)
        
        self.column_table = QTableWidget()
        self.column_table.setColumnCount(2)
        self.column_table.setHorizontalHeaderLabels(["Export", "Field Name"])
        self.column_table.horizontalHeader().setStretchLastSection(True)
        self.column_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        left_layout.addWidget(self.column_table)
        
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(10, 0, 0, 0)
        
        right_label = QLabel("Select notes to export (drag columns to reorder):")
        right_layout.addWidget(right_label)
        
        self.note_table = QTableWidget()
        self.note_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        
        self.note_table.horizontalHeader().setSectionsMovable(True)
        self.note_table.horizontalHeader().setDragEnabled(True)
        self.note_table.horizontalHeader().setDragDropMode(QTableWidget.DragDropMode.InternalMove)
        
        right_layout.addWidget(self.note_table)
        
        layout.addWidget(left_widget)
        layout.addWidget(right_widget)
        
        container = QWidget()
        container.setLayout(layout)
        
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(container)
        
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(scroll_area)
        
        selection_layout = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        select_all_btn.clicked.connect(self.selectAll)
        select_none_btn = QPushButton("Select None")
        select_none_btn.clicked.connect(self.selectNone)
        selection_layout.addWidget(select_all_btn)
        selection_layout.addWidget(select_none_btn)
        selection_layout.addStretch()
        main_layout.addLayout(selection_layout)
        
        audio_layout = QHBoxLayout()
        audio_label = QLabel("Audio folder name:")
        self.audio_edit = QLineEdit("audio")
        audio_layout.addWidget(audio_label)
        audio_layout.addWidget(self.audio_edit)
        main_layout.addLayout(audio_layout)
        
        format_layout = QHBoxLayout()
        format_label = QLabel("Export format:")
        self.format_combo = QComboBox()
        self.format_combo.addItems(["TSV", "CSV"])
        format_layout.addWidget(format_label)
        format_layout.addWidget(self.format_combo)
        main_layout.addLayout(format_layout)
        
        button_layout = QHBoxLayout()
        self.export_btn = QPushButton("Export Selected")
        self.export_btn.clicked.connect(self.onExport)
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.export_btn)
        button_layout.addWidget(self.cancel_btn)
        main_layout.addLayout(button_layout)
    
    def loadNoteInfo(self):
        field_names = set()
        self.notes_data = []
        
        for nid in self.nids:
            note = self.col.get_note(nid)
            note_data = {}
            for field_name in note.keys():
                field_names.add(field_name)
                note_data[field_name] = note[field_name] or ""
            self.notes_data.append((nid, note_data))
        
        self.field_order = sorted(field_names)
        
        self.column_table.setRowCount(len(self.field_order))
        for row, field_name in enumerate(self.field_order):
            checkbox = QCheckBox()
            checkbox.setChecked(True)
            checkbox_widget = QWidget()
            checkbox_layout = QHBoxLayout(checkbox_widget)
            checkbox_layout.addWidget(checkbox)
            checkbox_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            checkbox_layout.setContentsMargins(0, 0, 0, 0)
            self.column_table.setCellWidget(row, 0, checkbox_widget)
            
            self.column_table.setItem(row, 1, QTableWidgetItem(field_name))
        
        self.note_table.setRowCount(len(self.notes_data))
        
        headers = ["Export"] + self.field_order
        self.note_table.setColumnCount(len(headers))
        self.note_table.setHorizontalHeaderLabels(headers)
        
        for row, (nid, note_data) in enumerate(self.notes_data):
            checkbox = QCheckBox()
            checkbox.setChecked(True)
            checkbox_widget = QWidget()
            checkbox_layout = QHBoxLayout(checkbox_widget)
            checkbox_layout.addWidget(checkbox)
            checkbox_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            checkbox_layout.setContentsMargins(0, 0, 0, 0)
            self.note_table.setCellWidget(row, 0, checkbox_widget)
            
            for col, field_name in enumerate(self.field_order, 1):
                content = note_data.get(field_name, "")
                if len(content) > 100:
                    display_content = content[:100] + "..."
                else:
                    display_content = content
                
                item = QTableWidgetItem(display_content)
                item.setToolTip(content)
                self.note_table.setItem(row, col, item)
        
        self.column_table.resizeColumnsToContents()
        self.note_table.resizeColumnsToContents()
    
    def selectAll(self):
        for row in range(self.column_table.rowCount()):
            checkbox = self.column_table.cellWidget(row, 0).findChild(QCheckBox)
            checkbox.setChecked(True)
        
        for row in range(self.note_table.rowCount()):
            checkbox = self.note_table.cellWidget(row, 0).findChild(QCheckBox)
            checkbox.setChecked(True)
    
    def selectNone(self):
        for row in range(self.column_table.rowCount()):
            checkbox = self.column_table.cellWidget(row, 0).findChild(QCheckBox)
            checkbox.setChecked(False)
        
        for row in range(self.note_table.rowCount()):
            checkbox = self.note_table.cellWidget(row, 0).findChild(QCheckBox)
            checkbox.setChecked(False)
    
    def getVisualColumnOrder(self):
        visual_order = []
        for i in range(self.note_table.columnCount()):
            logical_index = self.note_table.horizontalHeader().logicalIndex(i)
            header_item = self.note_table.horizontalHeaderItem(logical_index)
            if header_item:
                visual_order.append(header_item.text())
        return visual_order
    
    def onExport(self):
        visual_column_order = self.getVisualColumnOrder()
        
        field_visual_order = visual_column_order[1:]
        
        selected_fields_from_checkbox = set()
        for row in range(self.column_table.rowCount()):
            checkbox = self.column_table.cellWidget(row, 0).findChild(QCheckBox)
            if checkbox.isChecked():
                field_name = self.column_table.item(row, 1).text()
                selected_fields_from_checkbox.add(field_name)
        
        selected_fields = [field for field in field_visual_order if field in selected_fields_from_checkbox]
    
        selected_rows = []
        for row in range(self.note_table.rowCount()):
            checkbox = self.note_table.cellWidget(row, 0).findChild(QCheckBox)
            if checkbox.isChecked():
                selected_rows.append(row)
    
        if not selected_fields:
            showInfo("Select at least one field to export.")
            return
    
        if not selected_rows:
            showInfo("Select at least one note to export.")
            return
    
        format_index = self.format_combo.currentIndex()
        if format_index == 0:
            file_filter = ".tsv"
            ext = ".tsv"
        else:
            file_filter = ".csv"
            ext = ".csv"
    
        path = getSaveFile(self, "Export notes", "notes_export", ext, file_filter)
        if not path:
            return
    
        audio_folder_name = self.audio_edit.text().strip() or "audio"
        export_dir = os.path.dirname(path)
        audio_dir = os.path.join(export_dir, audio_folder_name)
    
        if not os.path.exists(audio_dir):
            os.makedirs(audio_dir)
    
        try:
            delimiter = "\t" if format_index == 0 else ","
            self.exportText(path, selected_fields, selected_rows, audio_dir, audio_folder_name, delimiter)
            
            showInfo(f"Exported {len(selected_rows)} notes with {len(selected_fields)} fields successfully!\nAudio files saved to: {audio_dir}")
            self.accept()
        except Exception as e:
            showInfo(f"Error during export: {str(e)}")
    
    def exportText(self, path, fields, selected_rows, audio_dir, audio_folder_name, delimiter):
        with open(path, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f, delimiter=delimiter)
            writer.writerow(fields)
            
            for row_index in selected_rows:
                nid, note_data = self.notes_data[row_index]
                note = self.col.get_note(nid)
                row_data = []
                
                for field in fields:
                    content = ""
                    if field in note.keys():
                        content = note[field]
                    
                    if "[sound:" in content:
                        audio_files = []
                        start = content.find("[sound:")
                        while start != -1:
                            end = content.find("]", start)
                            if end == -1:
                                break
                            audio_file = content[start+7:end]
                            audio_files.append(audio_file)
                            start = content.find("[sound:", end)
                        
                        for audio_file in audio_files:
                            src_path = os.path.join(self.col.media.dir(), audio_file)
                            if os.path.exists(src_path):
                                dst_path = os.path.join(audio_dir, audio_file)
                                shutil.copy2(src_path, dst_path)
                                content = content.replace(f"[sound:{audio_file}]", f"{audio_folder_name}/{audio_file}")
                    
                    row_data.append(content)
                
                writer.writerow(row_data)

addHook("profileLoaded", setupMenu)
