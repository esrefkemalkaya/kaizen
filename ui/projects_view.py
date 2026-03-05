from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QDialog, QFormLayout, QLineEdit, QDialogButtonBox,
    QMessageBox, QLabel
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QColor

import db.models as m
from .styles import PAGE_TITLE_STYLE, TABLE_STYLE, BTN_PRIMARY, BTN_DANGER


class ProjectDialog(QDialog):
    def __init__(self, parent=None, name="", location=""):
        super().__init__(parent)
        self.setWindowTitle("Project")
        self.setMinimumWidth(380)
        layout = QFormLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        self.name_edit = QLineEdit(name)
        self.name_edit.setPlaceholderText("e.g. Project Alpha")
        self.loc_edit = QLineEdit(location)
        self.loc_edit.setPlaceholderText("e.g. Level 3, Block B")

        layout.addRow("Project Name *", self.name_edit)
        layout.addRow("Location", self.loc_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self._accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def _accept(self):
        if not self.name_edit.text().strip():
            QMessageBox.warning(self, "Validation", "Project name is required.")
            return
        self.accept()

    def values(self):
        return self.name_edit.text().strip(), self.loc_edit.text().strip()


class ProjectsView(QWidget):
    project_selected = pyqtSignal(int, str)   # (project_id, project_name)

    def __init__(self):
        super().__init__()
        self._selected_id = None
        self._build_ui()
        self._load()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(16)

        title = QLabel("Projects")
        title.setStyleSheet(PAGE_TITLE_STYLE)
        layout.addWidget(title)

        # Toolbar
        bar = QHBoxLayout()
        self.add_btn = QPushButton("+ Add Project")
        self.add_btn.setStyleSheet(BTN_PRIMARY)
        self.add_btn.clicked.connect(self._add)
        self.edit_btn = QPushButton("Edit")
        self.edit_btn.setStyleSheet(BTN_PRIMARY)
        self.edit_btn.clicked.connect(self._edit)
        self.del_btn = QPushButton("Delete")
        self.del_btn.setStyleSheet(BTN_DANGER)
        self.del_btn.clicked.connect(self._delete)
        bar.addWidget(self.add_btn)
        bar.addWidget(self.edit_btn)
        bar.addWidget(self.del_btn)
        bar.addStretch()
        layout.addLayout(bar)

        # Table
        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["ID", "Project Name", "Location"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(0, 50)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.itemSelectionChanged.connect(self._on_select)
        self.table.doubleClicked.connect(self._on_double_click)
        layout.addWidget(self.table)

        hint = QLabel("Select a project to work with it across all screens.")
        hint.setStyleSheet("color: #78909c; font-size: 12px;")
        layout.addWidget(hint)

    def _load(self):
        projects = m.get_projects()
        self.table.setRowCount(len(projects))
        for row, p in enumerate(projects):
            self.table.setItem(row, 0, QTableWidgetItem(str(p["id"])))
            self.table.setItem(row, 1, QTableWidgetItem(p["name"]))
            self.table.setItem(row, 2, QTableWidgetItem(p["location"] or ""))
            for col in range(3):
                if item := self.table.item(row, col):
                    item.setData(Qt.ItemDataRole.UserRole, p["id"])

    def _on_select(self):
        rows = self.table.selectedItems()
        if rows:
            self._selected_id = rows[0].data(Qt.ItemDataRole.UserRole)
            name = self.table.item(self.table.currentRow(), 1).text()
            self.project_selected.emit(self._selected_id, name)

    def _on_double_click(self):
        self._edit()

    def _add(self):
        dlg = ProjectDialog(self)
        if dlg.exec():
            name, loc = dlg.values()
            m.add_project(name, loc)
            self._load()

    def _edit(self):
        if not self._selected_id:
            QMessageBox.information(self, "Select", "Please select a project first.")
            return
        row = self.table.currentRow()
        name = self.table.item(row, 1).text()
        loc = self.table.item(row, 2).text()
        dlg = ProjectDialog(self, name, loc)
        if dlg.exec():
            new_name, new_loc = dlg.values()
            m.update_project(self._selected_id, new_name, new_loc)
            self._load()

    def _delete(self):
        if not self._selected_id:
            QMessageBox.information(self, "Select", "Please select a project first.")
            return
        reply = QMessageBox.question(
            self, "Confirm Delete",
            "Delete this project and ALL its data?\nThis cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            m.delete_project(self._selected_id)
            self._selected_id = None
            self._load()
