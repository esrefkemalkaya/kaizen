from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QLabel, QComboBox, QMessageBox, QSplitter,
    QFileDialog, QLineEdit, QStyledItemDelegate, QCompleter
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor
from datetime import datetime, timedelta

import db.models as m
from .styles import (PAGE_TITLE_STYLE, TABLE_STYLE, BTN_PRIMARY,
                     BTN_DANGER, BTN_SUCCESS, COMBO_STYLE, LABEL_MUTED)

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

STANDBY_TYPES = [
    "Patlatma", "Elektrik arızası", "Su kesintisi",
    "Topograf beklendi", "Servis ekibi beklendi", "Hava koşulları",
    "Bakım", "Diğer",
]


def _calc_hours(start: str, end: str) -> float:
    for fmt in ["%H:%M", "%H.%M", "%H:%M:%S"]:
        try:
            s = datetime.strptime(start.strip(), fmt)
            e = datetime.strptime(end.strip(), fmt)
            diff = e - s
            if diff.total_seconds() < 0:
                diff += timedelta(days=1)
            return round(diff.total_seconds() / 3600, 2)
        except ValueError:
            continue
    return 0.0


class AutoCompleteDelegate(QStyledItemDelegate):
    def __init__(self, items: list[str] | None = None, parent=None):
        super().__init__(parent)
        self._items: list[str] = items or []

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
        self._sb_hole_delegate = AutoCompleteDelegate()
        self._type_delegate = AutoCompleteDelegate(STANDBY_TYPES)
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

        # Cols: Hole ID | Start Date | End Date | Start Depth | End Depth | Meters | Amount
        self.bh_table = QTableWidget()
        self.bh_table.setStyleSheet(TABLE_STYLE)
        self.bh_table.setColumnCount(7)
        self.bh_table.setHorizontalHeaderLabels([
            "Hole ID", "Start Date", "End Date",
            "Start Depth (m)", "End Depth (m)", "Meters", "Amount"
        ])
        self.bh_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for col, w in [(1, 100), (2, 100), (3, 110), (4, 110), (5, 90), (6, 120)]:
            self.bh_table.setColumnWidth(col, w)
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

        # Cols: Date | Hole ID | Start | End | Type | Detail | Hours | Amount
        self.sb_table = QTableWidget()
        self.sb_table.setStyleSheet(TABLE_STYLE)
        self.sb_table.setColumnCount(8)
        self.sb_table.setHorizontalHeaderLabels([
            "Date", "Hole ID", "Start", "End",
            "Type", "Detail / Reason", "Hours", "Amount"
        ])
        self.sb_table.setColumnWidth(0, 90)
        self.sb_table.setColumnWidth(1, 130)
        self.sb_table.setColumnWidth(2, 65)
        self.sb_table.setColumnWidth(3, 65)
        self.sb_table.setColumnWidth(4, 140)
        self.sb_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        self.sb_table.setColumnWidth(6, 65)
        self.sb_table.setColumnWidth(7, 110)
        self.sb_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.sb_table.verticalHeader().setVisible(False)
        self.sb_table.setItemDelegateForColumn(1, self._sb_hole_delegate)
        self.sb_table.setItemDelegateForColumn(4, self._type_delegate)
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
        splitter.setSizes([380, 280])
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
            hole_ids = m.get_all_hole_ids(cid)
            self._hole_delegate.set_completions(hole_ids)
            all_holes = m.get_all_hole_ids_for_standby(cid)
            self._sb_hole_delegate.set_completions(all_holes)

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
            self._bh_append_row(
                e["id"], e["hole_id"],
                e["start_date"], e["end_date"],
                e["start_depth"], e["end_depth"],
                e["meters_drilled"], rate_m
            )
        self.bh_table.blockSignals(False)
        self._bh_update_totals()

    def _bh_append_row(self, db_id, hole_id, start_date, end_date,
                       start_depth, end_depth, meters, rate_m):
        row = self.bh_table.rowCount()
        self.bh_table.insertRow(row)
        self._borehole_ids.append(db_id)

        self.bh_table.setItem(row, 0, QTableWidgetItem(str(hole_id)))
        self.bh_table.setItem(row, 1, QTableWidgetItem(str(start_date)))
        self.bh_table.setItem(row, 2, QTableWidgetItem(str(end_date)))
        self.bh_table.setItem(row, 3, QTableWidgetItem(f"{start_depth:.2f}"))
        self.bh_table.setItem(row, 4, QTableWidgetItem(f"{end_depth:.2f}"))
        self.bh_table.setItem(row, 5, QTableWidgetItem(f"{meters:.2f}"))

        amt_item = QTableWidgetItem(f"${meters * rate_m:,.2f}")
        amt_item.setFlags(amt_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        amt_item.setForeground(QColor("#1b5e20"))
        self.bh_table.setItem(row, 6, amt_item)

    def _bh_add_row(self):
        if not self._contractor_id:
            QMessageBox.information(self, "Select", "Select a contractor first.")
            return
        rate_m, _ = self._get_rate()
        self.bh_table.blockSignals(True)
        self._bh_append_row(None, "", "", "", 0, 0, 0, rate_m)
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
            start_date = (self.bh_table.item(row, 1) or QTableWidgetItem("")).text().strip()
            end_date   = (self.bh_table.item(row, 2) or QTableWidgetItem("")).text().strip()
            try:
                start_depth = float((self.bh_table.item(row, 3) or QTableWidgetItem("0")).text())
                end_depth   = float((self.bh_table.item(row, 4) or QTableWidgetItem("0")).text())
                meters      = float((self.bh_table.item(row, 5) or QTableWidgetItem("0")).text())
            except ValueError:
                errors.append(f"Row {row + 1}: Invalid numeric value.")
                continue
            db_id = self._borehole_ids[row]
            new_id = m.upsert_drilling_entry(
                db_id, cid, month, year,
                hole_id, start_date, end_date, start_depth, end_depth, meters
            )
            self._borehole_ids[row] = new_id
        if errors:
            QMessageBox.warning(self, "Validation Errors", "\n".join(errors))
        else:
            QMessageBox.information(self, "Saved", "Boreholes saved successfully.")
        self._load_boreholes()

    def _bh_on_item_changed(self, item):
        row = item.row()
        col = item.column()
        if col in (3, 4):
            # depth changed → auto-calc meters
            try:
                s = float((self.bh_table.item(row, 3) or QTableWidgetItem("0")).text())
                e = float((self.bh_table.item(row, 4) or QTableWidgetItem("0")).text())
                meters = max(0.0, e - s)
            except ValueError:
                return
            self.bh_table.blockSignals(True)
            self.bh_table.setItem(row, 5, QTableWidgetItem(f"{meters:.2f}"))
            self._bh_set_amount(row, meters)
            self.bh_table.blockSignals(False)
            self._bh_update_totals()
        elif col == 5:
            # meters manually edited → update amount
            try:
                meters = float((self.bh_table.item(row, 5) or QTableWidgetItem("0")).text())
            except ValueError:
                return
            self.bh_table.blockSignals(True)
            self._bh_set_amount(row, meters)
            self.bh_table.blockSignals(False)
            self._bh_update_totals()

    def _bh_set_amount(self, row, meters):
        rate_m, _ = self._get_rate()
        amt = QTableWidgetItem(f"${meters * rate_m:,.2f}")
        amt.setFlags(amt.flags() & ~Qt.ItemFlag.ItemIsEditable)
        amt.setForeground(QColor("#1b5e20"))
        self.bh_table.setItem(row, 6, amt)

    def _bh_update_totals(self):
        rate_m, _ = self._get_rate()
        total_m = total_amt = 0.0
        for row in range(self.bh_table.rowCount()):
            try:
                meters = float((self.bh_table.item(row, 5) or QTableWidgetItem("0")).text())
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
                        if "hole" in val or "kuyu" in val:
                            col_hole = cell.column
                            header_row = cell.row
                        elif "meter" in val or "metre" in val:
                            col_meters = cell.column
                if header_row:
                    break
            if not header_row or not col_hole or not col_meters:
                QMessageBox.warning(self, "Import Error",
                    "Could not find required columns (Hole/Kuyu and Meter/Metre).")
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
                self._bh_append_row(None, hole_id, "", "", 0, 0, meters, rate_m)
                imported += 1
            self.bh_table.blockSignals(False)
            self._bh_update_totals()
            QMessageBox.information(self, "Import Complete",
                f"Imported {imported} rows. Click Save Boreholes to keep them.")
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
            self._sb_append_row(
                e["id"],
                e["entry_date"],
                e["hole_id"],
                e["start_time"],
                e["end_time"],
                e["standby_type"],
                e["description"],
                e["hours"],
                rate_s
            )
        self.sb_table.blockSignals(False)
        self._sb_update_totals()

    def _sb_append_row(self, db_id, entry_date, hole_id, start_time,
                       end_time, standby_type, description, hours, rate_s):
        row = self.sb_table.rowCount()
        self.sb_table.insertRow(row)
        self._standby_ids.append(db_id)

        self.sb_table.setItem(row, 0, QTableWidgetItem(str(entry_date)))
        self.sb_table.setItem(row, 1, QTableWidgetItem(str(hole_id)))
        self.sb_table.setItem(row, 2, QTableWidgetItem(str(start_time)))
        self.sb_table.setItem(row, 3, QTableWidgetItem(str(end_time)))
        self.sb_table.setItem(row, 4, QTableWidgetItem(str(standby_type)))
        self.sb_table.setItem(row, 5, QTableWidgetItem(str(description)))

        hrs_item = QTableWidgetItem(f"{hours:.2f}")
        hrs_item.setFlags(hrs_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.sb_table.setItem(row, 6, hrs_item)

        amt_item = QTableWidgetItem(f"${hours * rate_s:,.2f}")
        amt_item.setFlags(amt_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        amt_item.setForeground(QColor("#e65100"))
        self.sb_table.setItem(row, 7, amt_item)

    def _sb_add_row(self):
        if not self._contractor_id:
            QMessageBox.information(self, "Select", "Select a contractor first.")
            return
        _, rate_s = self._get_rate()
        self.sb_table.blockSignals(True)
        self._sb_append_row(None, "", "", "", "", "", "", 0, rate_s)
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
            entry_date   = (self.sb_table.item(row, 0) or QTableWidgetItem("")).text().strip()
            hole_id      = (self.sb_table.item(row, 1) or QTableWidgetItem("")).text().strip()
            start_time   = (self.sb_table.item(row, 2) or QTableWidgetItem("")).text().strip()
            end_time     = (self.sb_table.item(row, 3) or QTableWidgetItem("")).text().strip()
            standby_type = (self.sb_table.item(row, 4) or QTableWidgetItem("")).text().strip()
            description  = (self.sb_table.item(row, 5) or QTableWidgetItem("")).text().strip()

            if not start_time and not end_time and not description:
                errors.append(f"Row {row + 1}: Start/End time or description required.")
                continue

            hours = _calc_hours(start_time, end_time) if start_time and end_time else 0.0

            db_id = self._standby_ids[row]
            new_id = m.upsert_standby_entry(
                db_id, cid, month, year,
                entry_date, hole_id, start_time, end_time,
                standby_type, description, hours
            )
            self._standby_ids[row] = new_id

        if errors:
            QMessageBox.warning(self, "Validation Errors", "\n".join(errors))
        else:
            QMessageBox.information(self, "Saved", "Standby entries saved successfully.")
        self._load_standby()

    def _sb_on_item_changed(self, item):
        row = item.row()
        col = item.column()
        if col not in (2, 3):
            return
        start = (self.sb_table.item(row, 2) or QTableWidgetItem("")).text().strip()
        end   = (self.sb_table.item(row, 3) or QTableWidgetItem("")).text().strip()
        hours = _calc_hours(start, end) if start and end else 0.0
        _, rate_s = self._get_rate()

        self.sb_table.blockSignals(True)
        hrs_item = QTableWidgetItem(f"{hours:.2f}")
        hrs_item.setFlags(hrs_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.sb_table.setItem(row, 6, hrs_item)

        amt = QTableWidgetItem(f"${hours * rate_s:,.2f}")
        amt.setFlags(amt.flags() & ~Qt.ItemFlag.ItemIsEditable)
        amt.setForeground(QColor("#e65100"))
        self.sb_table.setItem(row, 7, amt)
        self.sb_table.blockSignals(False)
        self._sb_update_totals()

    def _sb_update_totals(self):
        _, rate_s = self._get_rate()
        total_h = total_amt = 0.0
        for row in range(self.sb_table.rowCount()):
            try:
                hours = float((self.sb_table.item(row, 6) or QTableWidgetItem("0")).text())
            except ValueError:
                continue
            total_h += hours
            total_amt += hours * rate_s
        self.sb_hours_lbl.setText(f"Total Hours: {total_h:,.2f}")
        self.sb_amount_lbl.setText(f"Total: ${total_amt:,.2f}")
