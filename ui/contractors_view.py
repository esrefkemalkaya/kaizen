from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QDialog, QFormLayout, QLineEdit, QComboBox,
    QDialogButtonBox, QMessageBox, QLabel, QDoubleSpinBox
)
from PyQt6.QtCore import Qt

import db.models as m
from .styles import PAGE_TITLE_STYLE, TABLE_STYLE, BTN_PRIMARY, BTN_DANGER, LABEL_MUTED


class ContractorDialog(QDialog):
    def __init__(self, parent=None, name="", ctype="underground", rate=0.0, standby=0.0):
        super().__init__(parent)
        self.setWindowTitle("Contractor")
        self.setMinimumWidth(400)
        layout = QFormLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        self.name_edit = QLineEdit(name)
        self.name_edit.setPlaceholderText("e.g. ABC Drilling Co.")

        self.type_combo = QComboBox()
        self.type_combo.addItem("Underground", "underground")
        self.type_combo.addItem("Surface", "surface")
        if ctype == "surface":
            self.type_combo.setCurrentIndex(1)

        self.rate_spin = QDoubleSpinBox()
        self.rate_spin.setRange(0, 999999)
        self.rate_spin.setDecimals(2)
        self.rate_spin.setPrefix("$ ")
        self.rate_spin.setValue(rate)

        self.standby_spin = QDoubleSpinBox()
        self.standby_spin.setRange(0, 999999)
        self.standby_spin.setDecimals(2)
        self.standby_spin.setPrefix("$ ")
        self.standby_spin.setValue(standby)

        layout.addRow("Contractor Name *", self.name_edit)
        layout.addRow("Type", self.type_combo)
        layout.addRow("Rate per Meter", self.rate_spin)
        layout.addRow("Standby Hour Rate", self.standby_spin)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self._accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def _accept(self):
        if not self.name_edit.text().strip():
            QMessageBox.warning(self, "Validation", "Contractor name is required.")
            return
        self.accept()

    def values(self):
        return (
            self.name_edit.text().strip(),
            self.type_combo.currentData(),
            self.rate_spin.value(),
            self.standby_spin.value(),
        )


class ContractorsView(QWidget):
    def __init__(self):
        super().__init__()
        self._project_id = None
        self._project_name = ""
        self._selected_id = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(16)

        self.title = QLabel("Contractors")
        self.title.setStyleSheet(PAGE_TITLE_STYLE)
        layout.addWidget(self.title)

        self.ctx_label = QLabel("No project selected — go to Projects and select one.")
        self.ctx_label.setStyleSheet(LABEL_MUTED)
        layout.addWidget(self.ctx_label)

        bar = QHBoxLayout()
        self.add_btn = QPushButton("+ Add Contractor")
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

        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["ID", "Name", "Type", "Rate/m", "Standby/hr"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(0, 50)
        self.table.setColumnWidth(2, 110)
        self.table.setColumnWidth(3, 110)
        self.table.setColumnWidth(4, 110)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.itemSelectionChanged.connect(self._on_select)
        self.table.doubleClicked.connect(self._on_double_click)
        layout.addWidget(self.table)

    def set_project(self, project_id: int, project_name: str):
        self._project_id = project_id
        self._project_name = project_name
        self.ctx_label.setText(f"Project: {project_name}")
        self._load()

    def _load(self):
        if not self._project_id:
            return
        rows = m.get_contractors(self._project_id)
        self.table.setRowCount(len(rows))
        for r, c in enumerate(rows):
            self.table.setItem(r, 0, QTableWidgetItem(str(c["id"])))
            self.table.setItem(r, 1, QTableWidgetItem(c["name"]))
            self.table.setItem(r, 2, QTableWidgetItem(c["type"].capitalize()))
            self.table.setItem(r, 3, QTableWidgetItem(f"${c['rate_per_meter']:.2f}"))
            self.table.setItem(r, 4, QTableWidgetItem(f"${c['standby_hour_rate']:.2f}"))
            for col in range(5):
                if item := self.table.item(r, col):
                    item.setData(Qt.ItemDataRole.UserRole, c["id"])

    def _on_select(self):
        rows = self.table.selectedItems()
        if rows:
            self._selected_id = rows[0].data(Qt.ItemDataRole.UserRole)

    def _on_double_click(self):
        self._edit()

    def _add(self):
        if not self._project_id:
            QMessageBox.information(self, "Select Project", "Please select a project first.")
            return
        dlg = ContractorDialog(self)
        if dlg.exec():
            name, ctype, rate, standby = dlg.values()
            m.add_contractor(self._project_id, name, ctype, rate, standby)
            self._load()

    def _edit(self):
        if not self._selected_id:
            QMessageBox.information(self, "Select", "Please select a contractor first.")
            return
        row = self.table.currentRow()
        name = self.table.item(row, 1).text()
        ctype = self.table.item(row, 2).text().lower()
        rate = float(self.table.item(row, 3).text().replace("$", ""))
        standby = float(self.table.item(row, 4).text().replace("$", ""))
        dlg = ContractorDialog(self, name, ctype, rate, standby)
        if dlg.exec():
            new_name, new_ctype, new_rate, new_standby = dlg.values()
            m.update_contractor(self._selected_id, new_name, new_ctype, new_rate, new_standby)
            self._load()

    def _delete(self):
        if not self._selected_id:
            QMessageBox.information(self, "Select", "Please select a contractor first.")
            return
        reply = QMessageBox.question(
            self, "Confirm Delete",
            "Delete this contractor and all their drilling/charge data?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            m.delete_contractor(self._selected_id)
            self._selected_id = None
            self._load()
