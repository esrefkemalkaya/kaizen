from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QLabel, QComboBox, QMessageBox, QSplitter,
    QFileDialog, QLineEdit, QStyledItemDelegate, QCompleter
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

import db.models as m
from .styles import (PAGE_TITLE_STYLE, TABLE_STYLE, BTN_PRIMARY,
                     BTN_DANGER, BTN_SUCCESS, COMBO_STYLE, LABEL_MUTED)

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]


class AutoCompleteDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._items: list[str] = []

    def set_completions(self, items: list[str]):
        self._items = items

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        completer = QCompleter(self._items, editor)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        completer.setFilterMode(Qt.MatchFlag.MatchContains)
        editor.setCompleter(completer)
        return editor

    def setEditorData(self, editor, index):
        editor.setText(index.data() or "")

    def setModelData(self, editor, model, index):
        model.setData(index, editor.text(), Qt.ItemDataRole.EditRole)


class DrillingEntryView(QWidget):
    def __init__(self):
        super().__init__()
        self._project_id = None
        self._project_name = ""
        self._contractor_id = None
        self._borehole_ids: list[int | None] = []
        self._standby_ids: list[int | None] = []
        self._hole_delegate = AutoCompleteDelegate()
        self._rig_delegate = AutoCompleteDelegate()
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(14)

        title = QLabel("Drilling Entries")
        title.setStyleSheet(PAGE_TITLE_STYLE)
        layout.addWidget(title)

        self.ctx_label = QLabel("No project selected.")
        self.ctx_label.setStyleSheet(LABEL_MUTED)
        layout.addWidget(self.ctx_label)

        # ── Filter bar ───────────────────────────────────────────────────
        filter_bar = QHBoxLayout()
        filter_bar.setSpacing(12)

        filter_bar.addWidget(QLabel("Contractor:"))
        self.contractor_combo = QComboBox()
        self.contractor_combo.setStyleSheet(COMBO_STYLE)
        self.contractor_combo.setMinimumWidth(200)
        self.contractor_combo.currentIndexChanged.connect(self._on_contractor_changed)
        filter_bar.addWidget(self.contractor_combo)

        filter_bar.addWidget(QLabel("Month:"))
        self.month_combo = QComboBox()
        self.month_combo.setStyleSheet(COMBO_STYLE)
        for i, name in enumerate(MONTHS):
            self.month_combo.addItem(name, i + 1)
        filter_bar.addWidget(self.month_combo)

        filter_bar.addWidget(QLabel("Year:"))
        self.year_combo = QComboBox()
        self.year_combo.setStyleSheet(COMBO_STYLE)
        from datetime import date
        current_year = date.today().year
        for y in range(current_year - 2, current_year + 3):
            self.year_combo.addItem(str(y), y)
        self.year_combo.setCurrentText(str(current_year))
        filter_bar.addWidget(self.year_combo)

        load_btn = QPushButton("Load")
        load_btn.setStyleSheet(BTN_PRIMARY)
        load_btn.clicked.connect(self._load_all)
        filter_bar.addWidget(load_btn)
        filter_bar.addStretch()
        layout.addLayout(filter_bar)

        # ── Splitter (Boreholes top / Standby bottom) ────────────────────
        splitter = QSplitter(Qt.Orientation.Vertical)

        # ── Boreholes section ─────────────────────────────────────────────
        bh_widget = QWidget()
        bh_layout = QVBoxLayout(bh_widget)
        bh_layout.setContentsMargins(0, 4, 0, 4)
        bh_layout.setSpacing(6)

        bh_label = QLabel("Boreholes")
        bh_label.setStyleSheet("font-weight: bold; font-size: 13px; color: #1565c0;")
        bh_layout.addWidget(bh_label)

        bh_action = QHBoxLayout()
        self.bh_add_btn = QPushButton("+ Add Row")
        self.bh_add_btn.setStyleSheet(BTN_PRIMARY)
        self.bh_add_btn.clicked.connect(self._bh_add_row)
        self.bh_del_btn = QPushButton("Delete Row")
        self.bh_del_btn.setStyleSheet(BTN_DANGER)
        self.bh_del_btn.clicked.connect(self._bh_delete_row)
        self.bh_save_btn = QPushButton("Save Boreholes")
        self.bh_save_btn.setStyleSheet(BTN_SUCCESS)
        self.bh_save_btn.clicked.connect(self._bh_save)
        self.bh_import_btn = QPushButton("Import from Excel")
        self.bh_import_btn.setStyleSheet(BTN_PRIMARY)
        self.bh_import_btn.clicked.connect(self._import_excel)
        for btn in [self.bh_add_btn, self.bh_del_btn, self.bh_save_btn, self.bh_import_btn]:
            bh_action.addWidget(btn)
        bh_action.addStretch()
        bh_layout.addLayout(bh_action)

        self.bh_table = QTableWidget()
        self.bh_table.setStyleSheet(TABLE_STYLE)
        self.bh_table.setColumnCount(3)
        self.bh_table.setHorizontalHeaderLabels(["Hole ID", "Meters Drilled", "Meter Amount"])
        self.bh_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.bh_table.setColumnWidth(1, 150)
        self.bh_table.setColumnWidth(2, 150)
        self.bh_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.bh_table.verticalHeader().setVisible(False)
        self.bh_table.setItemDelegateForColumn(0, self._hole_delegate)
        self.bh_table.itemChanged.connect(self._bh_on_item_changed)
        bh_layout.addWidget(self.bh_table)

        bh_footer = QHBoxLayout()
        self.bh_meters_lbl = QLabel("Total Meters: 0.00")
        self.bh_amount_lbl = QLabel("Total: $0.00")
        self.bh_amount_lbl.setStyleSheet("font-weight: bold; color: #1b5e20;")
        bh_footer.addWidget(self.bh_meters_lbl)
        bh_footer.addStretch()
        bh_footer.addWidget(self.bh_amount_lbl)
        bh_layout.addLayout(bh_footer)

        splitter.addWidget(bh_widget)

        # ── Standby section ───────────────────────────────────────────────
        sb_widget = QWidget()
        sb_layout = QVBoxLayout(sb_widget)
        sb_layout.setContentsMargins(0, 4, 0, 4)
        sb_layout.setSpacing(6)

        sb_label = QLabel("Standby Hours")
        sb_label.setStyleSheet("font-weight: bold; font-size: 13px; color: #e65100;")
        sb_layout.addWidget(sb_label)

        sb_action = QHBoxLayout()
        self.sb_add_btn = QPushButton("+ Add Row")
        self.sb_add_btn.setStyleSheet(BTN_PRIMARY)
        self.sb_add_btn.clicked.connect(self._sb_add_row)
        self.sb_del_btn = QPushButton("Delete Row")
        self.sb_del_btn.setStyleSheet(BTN_DANGER)
        self.sb_del_btn.clicked.connect(self._sb_delete_row)
        self.sb_save_btn = QPushButton("Save Standby")
        self.sb_save_btn.setStyleSheet(BTN_SUCCESS)
        self.sb_save_btn.clicked.connect(self._sb_save)
        for btn in [self.sb_add_btn, self.sb_del_btn, self.sb_save_btn]:
            sb_action.addWidget(btn)
        sb_action.addStretch()
        sb_layout.addLayout(sb_action)

        self.sb_table = QTableWidget()
        self.sb_table.setStyleSheet(TABLE_STYLE)
        self.sb_table.setColumnCount(4)
        self.sb_table.setHorizontalHeaderLabels(
            ["Rig Name", "Description / Reason", "Hours", "Amount"]
        )
        self.sb_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.sb_table.setColumnWidth(0, 160)
        self.sb_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.sb_table.setColumnWidth(2, 100)
        self.sb_table.setColumnWidth(3, 130)
        self.sb_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.sb_table.verticalHeader().setVisible(False)
        self.sb_table.setItemDelegateForColumn(0, self._rig_delegate)
        self.sb_table.itemChanged.connect(self._sb_on_item_changed)
        sb_layout.addWidget(self.sb_table)

        sb_footer = QHBoxLayout()
        self.sb_hours_lbl = QLabel("Total Hours: 0.00")
        self.sb_amount_lbl = QLabel("Total: $0.00")
        self.sb_amount_lbl.setStyleSheet("font-weight: bold; color: #e65100;")
        sb_footer.addWidget(self.sb_hours_lbl)
        sb_footer.addStretch()
        sb_footer.addWidget(self.sb_amount_lbl)
        sb_layout.addLayout(sb_footer)

        splitter.addWidget(sb_widget)
        splitter.setSizes([380, 260])
        layout.addWidget(splitter)

    # ── Public API ────────────────────────────────────────────────────────────

    def set_project(self, project_id: int, project_name: str):
        self._project_id = project_id
        self._project_name = project_name
        self.ctx_label.setText(f"Project: {project_name}")
        self._refresh_contractors()

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _refresh_contractors(self):
        self.contractor_combo.blockSignals(True)
        self.contractor_combo.clear()
        if self._project_id:
            for c in m.get_contractors(self._project_id):
                self.contractor_combo.addItem(
                    f"{c['name']} ({c['type'].capitalize()})", c["id"]
                )
        self.contractor_combo.blockSignals(False)
        self._load_all()

    def _on_contractor_changed(self):
        self._contractor_id = self.contractor_combo.currentData()

    def _get_rate(self):
        cid = self.contractor_combo.currentData()
        if cid:
            c = m.get_contractor(cid)
            if c:
                return c["rate_per_meter"], c["standby_hour_rate"]
        return 0.0, 0.0

    def _refresh_completions(self):
        cid = self.contractor_combo.currentData()
        if cid:
            self._hole_delegate.set_completions(m.get_all_hole_ids(cid))
            self._rig_delegate.set_completions(m.get_all_rig_names(cid))

    def _load_all(self):
        self._load_boreholes()
        self._load_standby()

    # ── Borehole table ────────────────────────────────────────────────────────

    def _load_boreholes(self):
        cid = self.contractor_combo.currentData()
        month = self.month_combo.currentData()
        year = self.year_combo.currentData()
        if not cid or not month or not year:
            return
        self._contractor_id = cid
        self._refresh_completions()
        entries = m.get_drilling_entries(cid, month, year)
        rate_m, _ = self._get_rate()
        self._borehole_ids = []
        self.bh_table.blockSignals(True)
        self.bh_table.setRowCount(0)
        for e in entries:
            self._bh_append_row(e["id"], e["hole_id"], e["meters_drilled"], rate_m)
        self.bh_table.blockSignals(False)
        self._bh_update_totals()

    def _bh_append_row(self, db_id, hole_id, meters, rate_m):
        row = self.bh_table.rowCount()
        self.bh_table.insertRow(row)
        self._borehole_ids.append(db_id)
        self.bh_table.setItem(row, 0, QTableWidgetItem(str(hole_id)))
        self.bh_table.setItem(row, 1, QTableWidgetItem(f"{meters:.2f}"))
        amt_item = QTableWidgetItem(f"${meters * rate_m:,.2f}")
        amt_item.setFlags(amt_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        amt_item.setForeground(QColor("#1b5e20"))
        self.bh_table.setItem(row, 2, amt_item)

    def _bh_add_row(self):
        if not self._contractor_id:
            QMessageBox.information(self, "Select", "Select a contractor first.")
            return
        rate_m, _ = self._get_rate()
        self.bh_table.blockSignals(True)
        self._bh_append_row(None, "", 0, rate_m)
        self.bh_table.blockSignals(False)

    def _bh_delete_row(self):
        row = self.bh_table.currentRow()
        if row < 0:
            return
        db_id = self._borehole_ids[row]
        if db_id:
            reply = QMessageBox.question(
                self, "Confirm", "Delete this row?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
            m.delete_drilling_entry(db_id)
        self.bh_table.removeRow(row)
        self._borehole_ids.pop(row)
        self._bh_update_totals()

    def _bh_save(self):
        cid = self.contractor_combo.currentData()
        month = self.month_combo.currentData()
        year = self.year_combo.currentData()
        if not cid:
            QMessageBox.information(self, "Select", "Select a contractor first.")
            return
        errors = []
        for row in range(self.bh_table.rowCount()):
            hole_id = (self.bh_table.item(row, 0) or QTableWidgetItem("")).text().strip()
            if not hole_id:
                errors.append(f"Row {row + 1}: Hole ID is empty.")
                continue
            try:
                meters = float((self.bh_table.item(row, 1) or QTableWidgetItem("0")).text())
            except ValueError:
                errors.append(f"Row {row + 1}: Invalid meters value.")
                continue
            db_id = self._borehole_ids[row]
            new_id = m.upsert_drilling_entry(db_id, cid, month, year, hole_id, meters)
            self._borehole_ids[row] = new_id
        if errors:
            QMessageBox.warning(self, "Validation Errors", "\n".join(errors))
        else:
            QMessageBox.information(self, "Saved", "Boreholes saved successfully.")
        self._load_boreholes()

    def _bh_on_item_changed(self, item):
        if item.column() != 1:
            return
        row = item.row()
        rate_m, _ = self._get_rate()
        try:
            meters = float((self.bh_table.item(row, 1) or QTableWidgetItem("0")).text())
        except ValueError:
            return
        self.bh_table.blockSignals(True)
        amt = QTableWidgetItem(f"${meters * rate_m:,.2f}")
        amt.setFlags(amt.flags() & ~Qt.ItemFlag.ItemIsEditable)
        amt.setForeground(QColor("#1b5e20"))
        self.bh_table.setItem(row, 2, amt)
        self.bh_table.blockSignals(False)
        self._bh_update_totals()

    def _bh_update_totals(self):
        rate_m, _ = self._get_rate()
        total_m = total_amt = 0.0
        for row in range(self.bh_table.rowCount()):
            try:
                meters = float((self.bh_table.item(row, 1) or QTableWidgetItem("0")).text())
            except ValueError:
                continue
            total_m += meters
            total_amt += meters * rate_m
        self.bh_meters_lbl.setText(f"Total Meters: {total_m:,.2f}")
        self.bh_amount_lbl.setText(f"Total: ${total_amt:,.2f}")

    def _import_excel(self):
        if not self._contractor_id:
            QMessageBox.information(self, "Select", "Select a contractor first.")
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Import Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if not path:
            return
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            header_row = col_hole = col_meters = None
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        val = cell.value.strip().lower()
                        if "hole" in val:
                            col_hole = cell.column
                            header_row = cell.row
                        elif "meter" in val:
                            col_meters = cell.column
                if header_row:
                    break
            if not header_row or not col_hole or not col_meters:
                QMessageBox.warning(
                    self, "Import Error",
                    "Could not find required columns.\n"
                    "The file must have columns containing 'Hole' and 'Meter'."
                )
                return
            rate_m, _ = self._get_rate()
            imported = 0
            self.bh_table.blockSignals(True)
            for row in ws.iter_rows(min_row=header_row + 1):
                hole_id = row[col_hole - 1].value
                meters_val = row[col_meters - 1].value
                if hole_id is None and meters_val is None:
                    continue
                hole_id = str(hole_id).strip() if hole_id is not None else ""
                try:
                    meters = float(meters_val) if meters_val is not None else 0.0
                except (ValueError, TypeError):
                    continue
                self._bh_append_row(None, hole_id, meters, rate_m)
                imported += 1
            self.bh_table.blockSignals(False)
            self._bh_update_totals()
            QMessageBox.information(
                self, "Import Complete",
                f"Imported {imported} rows. Click Save Boreholes to keep them."
            )
        except Exception as e:
            QMessageBox.critical(self, "Import Failed", str(e))

    # ── Standby table ─────────────────────────────────────────────────────────

    def _load_standby(self):
        cid = self.contractor_combo.currentData()
        month = self.month_combo.currentData()
        year = self.year_combo.currentData()
        if not cid or not month or not year:
            return
        _, rate_s = self._get_rate()
        entries = m.get_standby_entries(cid, month, year)
        self._standby_ids = []
        self.sb_table.blockSignals(True)
        self.sb_table.setRowCount(0)
        for e in entries:
            self._sb_append_row(e["id"], e["rig_name"], e["description"], e["hours"], rate_s)
        self.sb_table.blockSignals(False)
        self._sb_update_totals()

    def _sb_append_row(self, db_id, rig_name, description, hours, rate_s):
        row = self.sb_table.rowCount()
        self.sb_table.insertRow(row)
        self._standby_ids.append(db_id)
        self.sb_table.setItem(row, 0, QTableWidgetItem(str(rig_name)))
        self.sb_table.setItem(row, 1, QTableWidgetItem(str(description)))
        self.sb_table.setItem(row, 2, QTableWidgetItem(f"{hours:.2f}"))
        amt_item = QTableWidgetItem(f"${hours * rate_s:,.2f}")
        amt_item.setFlags(amt_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        amt_item.setForeground(QColor("#e65100"))
        self.sb_table.setItem(row, 3, amt_item)

    def _sb_add_row(self):
        if not self._contractor_id:
            QMessageBox.information(self, "Select", "Select a contractor first.")
            return
        _, rate_s = self._get_rate()
        self.sb_table.blockSignals(True)
        self._sb_append_row(None, "", "", 0, rate_s)
        self.sb_table.blockSignals(False)

    def _sb_delete_row(self):
        row = self.sb_table.currentRow()
        if row < 0:
            return
        db_id = self._standby_ids[row]
        if db_id:
            reply = QMessageBox.question(
                self, "Confirm", "Delete this row?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
            m.delete_standby_entry(db_id)
        self.sb_table.removeRow(row)
        self._standby_ids.pop(row)
        self._sb_update_totals()

    def _sb_save(self):
        cid = self.contractor_combo.currentData()
        month = self.month_combo.currentData()
        year = self.year_combo.currentData()
        if not cid:
            QMessageBox.information(self, "Select", "Select a contractor first.")
            return
        errors = []
        for row in range(self.sb_table.rowCount()):
            rig = (self.sb_table.item(row, 0) or QTableWidgetItem("")).text().strip()
            if not rig:
                errors.append(f"Row {row + 1}: Rig Name is empty.")
                continue
            desc = (self.sb_table.item(row, 1) or QTableWidgetItem("")).text().strip()
            try:
                hours = float((self.sb_table.item(row, 2) or QTableWidgetItem("0")).text())
            except ValueError:
                errors.append(f"Row {row + 1}: Invalid hours value.")
                continue
            db_id = self._standby_ids[row]
            new_id = m.upsert_standby_entry(db_id, cid, month, year, rig, desc, hours)
            self._standby_ids[row] = new_id
        if errors:
            QMessageBox.warning(self, "Validation Errors", "\n".join(errors))
        else:
            QMessageBox.information(self, "Saved", "Standby entries saved successfully.")
        self._load_standby()

    def _sb_on_item_changed(self, item):
        if item.column() != 2:
            return
        row = item.row()
        _, rate_s = self._get_rate()
        try:
            hours = float((self.sb_table.item(row, 2) or QTableWidgetItem("0")).text())
        except ValueError:
            return
        self.sb_table.blockSignals(True)
        amt = QTableWidgetItem(f"${hours * rate_s:,.2f}")
        amt.setFlags(amt.flags() & ~Qt.ItemFlag.ItemIsEditable)
        amt.setForeground(QColor("#e65100"))
        self.sb_table.setItem(row, 3, amt)
        self.sb_table.blockSignals(False)
        self._sb_update_totals()

    def _sb_update_totals(self):
        _, rate_s = self._get_rate()
        total_h = total_amt = 0.0
        for row in range(self.sb_table.rowCount()):
            try:
                hours = float((self.sb_table.item(row, 2) or QTableWidgetItem("0")).text())
            except ValueError:
                continue
            total_h += hours
            total_amt += hours * rate_s
        self.sb_hours_lbl.setText(f"Total Hours: {total_h:,.2f}")
        self.sb_amount_lbl.setText(f"Total: ${total_amt:,.2f}")
