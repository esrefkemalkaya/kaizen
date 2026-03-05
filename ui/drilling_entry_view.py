from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QLabel, QComboBox, QMessageBox, QAbstractItemView,
    QFileDialog, QLineEdit, QStyledItemDelegate, QCompleter
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QFont

import db.models as m
from .styles import (PAGE_TITLE_STYLE, TABLE_STYLE, BTN_PRIMARY,
                     BTN_DANGER, BTN_SUCCESS, COMBO_STYLE, LABEL_MUTED)

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]


class HoleIdDelegate(QStyledItemDelegate):
    """Delegate for col 0 that shows a QLineEdit with autocomplete."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._hole_ids: list[str] = []

    def set_completions(self, hole_ids: list[str]):
        self._hole_ids = hole_ids

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        completer = QCompleter(self._hole_ids, editor)
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
        self._row_ids: list[int | None] = []   # DB id per table row
        self._hole_id_delegate = HoleIdDelegate()
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
        load_btn.clicked.connect(self._load_entries)
        filter_bar.addWidget(load_btn)
        filter_bar.addStretch()
        layout.addLayout(filter_bar)

        # ── Action bar ───────────────────────────────────────────────────
        action_bar = QHBoxLayout()
        self.add_row_btn = QPushButton("+ Add Row")
        self.add_row_btn.setStyleSheet(BTN_PRIMARY)
        self.add_row_btn.clicked.connect(self._add_row)

        self.del_row_btn = QPushButton("Delete Row")
        self.del_row_btn.setStyleSheet(BTN_DANGER)
        self.del_row_btn.clicked.connect(self._delete_row)

        self.save_btn = QPushButton("Save All")
        self.save_btn.setStyleSheet(BTN_SUCCESS)
        self.save_btn.clicked.connect(self._save_all)

        self.import_btn = QPushButton("Import from Excel")
        self.import_btn.setStyleSheet(BTN_PRIMARY)
        self.import_btn.clicked.connect(self._import_excel)

        action_bar.addWidget(self.add_row_btn)
        action_bar.addWidget(self.del_row_btn)
        action_bar.addWidget(self.save_btn)
        action_bar.addWidget(self.import_btn)
        action_bar.addStretch()
        layout.addLayout(action_bar)

        # ── Table ────────────────────────────────────────────────────────
        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE)
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "Hole ID", "Meters Drilled", "Standby Hours",
            "Meter Amount", "Standby Amount", "Total"
        ])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for col in range(1, 6):
            self.table.setColumnWidth(col, 130)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.setItemDelegateForColumn(0, self._hole_id_delegate)
        self.table.itemChanged.connect(self._on_item_changed)
        layout.addWidget(self.table)

        # ── Totals footer ────────────────────────────────────────────────
        footer = QHBoxLayout()
        self.total_meters_lbl = QLabel("Total Meters: 0.00")
        self.total_standby_lbl = QLabel("Total Standby Hrs: 0.00")
        self.total_amount_lbl = QLabel("Grand Total: $0.00")
        self.total_amount_lbl.setStyleSheet("font-weight: bold; font-size: 14px; color: #1b5e20;")
        for lbl in [self.total_meters_lbl, self.total_standby_lbl, self.total_amount_lbl]:
            lbl.setStyleSheet(getattr(lbl, 'styleSheet', lambda: '')() +
                              " padding: 4px 16px;")
        footer.addWidget(self.total_meters_lbl)
        footer.addWidget(self.total_standby_lbl)
        footer.addStretch()
        footer.addWidget(self.total_amount_lbl)
        layout.addLayout(footer)

    # ── Public API ────────────────────────────────────────────────────────────

    def set_project(self, project_id: int, project_name: str):
        self._project_id = project_id
        self._project_name = project_name
        self.ctx_label.setText(f"Project: {project_name}")
        self._refresh_contractors()

    # ── Internal helpers ──────────────────────────────────────────────────────

    def _refresh_contractors(self):
        self.contractor_combo.blockSignals(True)
        self.contractor_combo.clear()
        if self._project_id:
            for c in m.get_contractors(self._project_id):
                self.contractor_combo.addItem(
                    f"{c['name']} ({c['type'].capitalize()})", c["id"]
                )
        self.contractor_combo.blockSignals(False)
        self._load_entries()

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
            self._hole_id_delegate.set_completions(hole_ids)

    def _load_entries(self):
        cid = self.contractor_combo.currentData()
        month = self.month_combo.currentData()
        year = self.year_combo.currentData()
        if not cid or not month or not year:
            return
        self._contractor_id = cid
        self._refresh_completions()
        entries = m.get_drilling_entries(cid, month, year)
        self._row_ids = []
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        rate_m, rate_s = self._get_rate()
        for e in entries:
            self._append_row(
                e["id"], e["hole_id"],
                e["meters_drilled"], e["standby_hours"],
                rate_m, rate_s
            )
        self.table.blockSignals(False)
        self._update_totals()

    def _append_row(self, db_id, hole_id, meters, standby, rate_m, rate_s):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self._row_ids.append(db_id)

        meter_amt = meters * rate_m
        standby_amt = standby * rate_s
        total = meter_amt + standby_amt

        self.table.setItem(row, 0, QTableWidgetItem(str(hole_id)))
        self.table.setItem(row, 1, QTableWidgetItem(f"{meters:.2f}"))
        self.table.setItem(row, 2, QTableWidgetItem(f"{standby:.2f}"))

        for col, val in [(3, meter_amt), (4, standby_amt), (5, total)]:
            item = QTableWidgetItem(f"${val:,.2f}")
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            item.setForeground(QColor("#1b5e20"))
            self.table.setItem(row, col, item)

    def _add_row(self):
        if not self._contractor_id:
            QMessageBox.information(self, "Select", "Select a contractor first.")
            return
        rate_m, rate_s = self._get_rate()
        self.table.blockSignals(True)
        self._append_row(None, "", 0, 0, rate_m, rate_s)
        self.table.blockSignals(False)

    def _delete_row(self):
        row = self.table.currentRow()
        if row < 0:
            return
        db_id = self._row_ids[row]
        if db_id:
            reply = QMessageBox.question(
                self, "Confirm", "Delete this row?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
            m.delete_drilling_entry(db_id)
        self.table.removeRow(row)
        self._row_ids.pop(row)
        self._update_totals()

    def _save_all(self):
        cid = self.contractor_combo.currentData()
        month = self.month_combo.currentData()
        year = self.year_combo.currentData()
        if not cid:
            QMessageBox.information(self, "Select", "Select a contractor first.")
            return

        errors = []
        for row in range(self.table.rowCount()):
            hole_id = (self.table.item(row, 0) or QTableWidgetItem("")).text().strip()
            if not hole_id:
                errors.append(f"Row {row + 1}: Hole ID is empty.")
                continue
            try:
                meters = float((self.table.item(row, 1) or QTableWidgetItem("0")).text())
                standby = float((self.table.item(row, 2) or QTableWidgetItem("0")).text())
            except ValueError:
                errors.append(f"Row {row + 1}: Invalid number.")
                continue
            db_id = self._row_ids[row]
            new_id = m.upsert_drilling_entry(db_id, cid, month, year, hole_id, meters, standby)
            self._row_ids[row] = new_id

        if errors:
            QMessageBox.warning(self, "Validation Errors", "\n".join(errors))
        else:
            QMessageBox.information(self, "Saved", "All entries saved successfully.")

        self._load_entries()

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

            # Find header row — look for Hole ID, Meters, Standby columns
            header_row = None
            col_hole = col_meters = col_standby = None
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        val = cell.value.strip().lower()
                        if "hole" in val:
                            col_hole = cell.column
                            header_row = cell.row
                        elif "meter" in val:
                            col_meters = cell.column
                        elif "standby" in val:
                            col_standby = cell.column
                if header_row:
                    break

            if not header_row or not col_hole or not col_meters:
                QMessageBox.warning(
                    self, "Import Error",
                    "Could not find required columns.\n"
                    "The file must have columns containing 'Hole', 'Meter', and optionally 'Standby'."
                )
                return

            rate_m, rate_s = self._get_rate()
            imported = 0
            self.table.blockSignals(True)
            for row in ws.iter_rows(min_row=header_row + 1):
                hole_id = row[col_hole - 1].value
                meters_val = row[col_meters - 1].value
                standby_val = row[col_standby - 1].value if col_standby else 0

                if hole_id is None and meters_val is None:
                    continue  # skip blank rows

                hole_id = str(hole_id).strip() if hole_id is not None else ""
                try:
                    meters = float(meters_val) if meters_val is not None else 0.0
                    standby = float(standby_val) if standby_val is not None else 0.0
                except (ValueError, TypeError):
                    continue

                self._append_row(None, hole_id, meters, standby, rate_m, rate_s)
                imported += 1

            self.table.blockSignals(False)
            self._update_totals()

            QMessageBox.information(
                self, "Import Complete",
                f"Imported {imported} rows. Review the data and click Save All to keep it."
            )

        except Exception as e:
            QMessageBox.critical(self, "Import Failed", str(e))

    def _on_item_changed(self, item):
        row = item.row()
        col = item.column()
        if col not in (1, 2):
            return
        rate_m, rate_s = self._get_rate()
        try:
            meters = float((self.table.item(row, 1) or QTableWidgetItem("0")).text())
            standby = float((self.table.item(row, 2) or QTableWidgetItem("0")).text())
        except ValueError:
            return
        meter_amt = meters * rate_m
        standby_amt = standby * rate_s
        total = meter_amt + standby_amt

        self.table.blockSignals(True)
        for c, val in [(3, meter_amt), (4, standby_amt), (5, total)]:
            it = QTableWidgetItem(f"${val:,.2f}")
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            it.setForeground(QColor("#1b5e20"))
            self.table.setItem(row, c, it)
        self.table.blockSignals(False)
        self._update_totals()

    def _update_totals(self):
        total_m = total_s = grand = 0.0
        rate_m, rate_s = self._get_rate()
        for row in range(self.table.rowCount()):
            try:
                meters = float((self.table.item(row, 1) or QTableWidgetItem("0")).text())
                standby = float((self.table.item(row, 2) or QTableWidgetItem("0")).text())
            except ValueError:
                continue
            total_m += meters
            total_s += standby
            grand += meters * rate_m + standby * rate_s

        self.total_meters_lbl.setText(f"Total Meters: {total_m:,.2f}")
        self.total_standby_lbl.setText(f"Total Standby Hrs: {total_s:,.2f}")
        self.total_amount_lbl.setText(f"Grand Total: ${grand:,.2f}")
