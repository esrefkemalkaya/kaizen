from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QLabel, QComboBox, QMessageBox, QTabWidget
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

import db.models as m
from .styles import (PAGE_TITLE_STYLE, TABLE_STYLE, BTN_PRIMARY,
                     BTN_DANGER, BTN_SUCCESS, COMBO_STYLE, LABEL_MUTED)

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]


class ChargesView(QWidget):
    def __init__(self):
        super().__init__()
        self._project_id = None
        self._project_name = ""
        self._ppe_row_ids: list[int | None] = []
        self._diesel_row_ids: list[int | None] = []
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(14)

        title = QLabel("Charges (PPE & Diesel)")
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
        load_btn.clicked.connect(self._load_charges)
        filter_bar.addWidget(load_btn)
        filter_bar.addStretch()
        layout.addLayout(filter_bar)

        # ── Tabs ─────────────────────────────────────────────────────────
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabBar::tab { padding: 8px 24px; font-size: 13px; }
            QTabBar::tab:selected { background: #1565c0; color: white; font-weight: bold; }
        """)

        self.ppe_widget = self._build_charge_table(
            is_ppe=True,
            col_label="Item Name",
            placeholder="e.g. Safety Helmet"
        )
        self.diesel_widget = self._build_charge_table(
            is_ppe=False,
            col_label="Description",
            placeholder="e.g. Diesel for generator"
        )

        self.tabs.addTab(self.ppe_widget, "PPE Charges (Underground)")
        self.tabs.addTab(self.diesel_widget, "Diesel Charges (Surface)")
        layout.addWidget(self.tabs)

    def _build_charge_table(self, is_ppe: bool, col_label: str, placeholder: str) -> QWidget:
        container = QWidget()
        v = QVBoxLayout(container)
        v.setContentsMargins(0, 12, 0, 0)
        v.setSpacing(10)

        bar = QHBoxLayout()
        add_btn = QPushButton("+ Add Row")
        add_btn.setStyleSheet(BTN_PRIMARY)
        del_btn = QPushButton("Delete Row")
        del_btn.setStyleSheet(BTN_DANGER)
        save_btn = QPushButton("Save All")
        save_btn.setStyleSheet(BTN_SUCCESS)
        bar.addWidget(add_btn)
        bar.addWidget(del_btn)
        bar.addWidget(save_btn)
        bar.addStretch()
        v.addLayout(bar)

        table = QTableWidget()
        table.setStyleSheet(TABLE_STYLE)
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels([col_label, "Quantity", "Unit Price", "Total"])
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        table.setColumnWidth(1, 110)
        table.setColumnWidth(2, 130)
        table.setColumnWidth(3, 130)
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        table.verticalHeader().setVisible(False)
        v.addWidget(table)

        # Footer total
        footer = QHBoxLayout()
        total_lbl = QLabel("Total: $0.00")
        total_lbl.setStyleSheet("font-weight: bold; font-size: 14px; color: #c62828; padding: 4px 0;")
        footer.addStretch()
        footer.addWidget(total_lbl)
        v.addLayout(footer)

        if is_ppe:
            self.ppe_table = table
            self.ppe_total_lbl = total_lbl
            add_btn.clicked.connect(self._add_ppe_row)
            del_btn.clicked.connect(self._delete_ppe_row)
            save_btn.clicked.connect(self._save_ppe)
            table.itemChanged.connect(lambda item: self._recalc_table(table, total_lbl))
        else:
            self.diesel_table = table
            self.diesel_total_lbl = total_lbl
            add_btn.clicked.connect(self._add_diesel_row)
            del_btn.clicked.connect(self._delete_diesel_row)
            save_btn.clicked.connect(self._save_diesel)
            table.itemChanged.connect(lambda item: self._recalc_table(table, total_lbl))

        return container

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
        self._load_charges()

    def _on_contractor_changed(self):
        pass  # charges loaded on demand via Load button

    def _load_charges(self):
        cid = self.contractor_combo.currentData()
        month = self.month_combo.currentData()
        year = self.year_combo.currentData()
        if not cid or not month or not year:
            return

        contractor = m.get_contractor(cid)
        if not contractor:
            return

        ctype = contractor["type"]

        # Load PPE (underground only)
        self._ppe_row_ids = []
        self.ppe_table.blockSignals(True)
        self.ppe_table.setRowCount(0)
        if ctype == "underground":
            for row in m.get_ppe_charges(cid, month, year):
                self._append_charge_row(
                    self.ppe_table, self._ppe_row_ids,
                    row["id"], row["item_name"], row["quantity"], row["unit_price"]
                )
            self.tabs.setTabEnabled(0, True)
        else:
            self.tabs.setTabEnabled(0, False)
        self.ppe_table.blockSignals(False)
        self._recalc_table(self.ppe_table, self.ppe_total_lbl)

        # Load Diesel (surface only)
        self._diesel_row_ids = []
        self.diesel_table.blockSignals(True)
        self.diesel_table.setRowCount(0)
        if ctype == "surface":
            for row in m.get_diesel_charges(cid, month, year):
                self._append_charge_row(
                    self.diesel_table, self._diesel_row_ids,
                    row["id"], row["description"], row["quantity"], row["unit_price"]
                )
            self.tabs.setTabEnabled(1, True)
        else:
            self.tabs.setTabEnabled(1, False)
        self.diesel_table.blockSignals(False)
        self._recalc_table(self.diesel_table, self.diesel_total_lbl)

        # Switch to applicable tab
        if ctype == "underground":
            self.tabs.setCurrentIndex(0)
        else:
            self.tabs.setCurrentIndex(1)

    def _append_charge_row(self, table, row_ids, db_id, name, qty, unit_price):
        row = table.rowCount()
        table.insertRow(row)
        row_ids.append(db_id)
        table.setItem(row, 0, QTableWidgetItem(str(name)))
        table.setItem(row, 1, QTableWidgetItem(f"{qty:.2f}"))
        table.setItem(row, 2, QTableWidgetItem(f"{unit_price:.2f}"))
        total = qty * unit_price
        it = QTableWidgetItem(f"${total:,.2f}")
        it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
        it.setForeground(QColor("#c62828"))
        table.setItem(row, 3, it)

    def _recalc_table(self, table: QTableWidget, total_lbl: QLabel):
        total = 0.0
        table.blockSignals(True)
        for row in range(table.rowCount()):
            try:
                qty = float((table.item(row, 1) or QTableWidgetItem("0")).text())
                price = float((table.item(row, 2) or QTableWidgetItem("0")).text())
            except ValueError:
                qty = price = 0.0
            val = qty * price
            total += val
            it = QTableWidgetItem(f"${val:,.2f}")
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            it.setForeground(QColor("#c62828"))
            table.setItem(row, 3, it)
        table.blockSignals(False)
        total_lbl.setText(f"Total: ${total:,.2f}")

    # ── PPE helpers ───────────────────────────────────────────────────────────

    def _add_ppe_row(self):
        if not self.contractor_combo.currentData():
            return
        self.ppe_table.blockSignals(True)
        self._append_charge_row(self.ppe_table, self._ppe_row_ids, None, "", 0, 0)
        self.ppe_table.blockSignals(False)

    def _delete_ppe_row(self):
        row = self.ppe_table.currentRow()
        if row < 0:
            return
        db_id = self._ppe_row_ids[row]
        if db_id:
            reply = QMessageBox.question(self, "Confirm", "Delete this row?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply != QMessageBox.StandardButton.Yes:
                return
            m.delete_ppe_charge(db_id)
        self.ppe_table.removeRow(row)
        self._ppe_row_ids.pop(row)
        self._recalc_table(self.ppe_table, self.ppe_total_lbl)

    def _save_ppe(self):
        cid = self.contractor_combo.currentData()
        month = self.month_combo.currentData()
        year = self.year_combo.currentData()
        if not cid:
            return
        errors = []
        for row in range(self.ppe_table.rowCount()):
            name = (self.ppe_table.item(row, 0) or QTableWidgetItem("")).text().strip()
            if not name:
                errors.append(f"Row {row + 1}: Item name is empty.")
                continue
            try:
                qty = float((self.ppe_table.item(row, 1) or QTableWidgetItem("0")).text())
                price = float((self.ppe_table.item(row, 2) or QTableWidgetItem("0")).text())
            except ValueError:
                errors.append(f"Row {row + 1}: Invalid number.")
                continue
            db_id = self._ppe_row_ids[row]
            new_id = m.upsert_ppe_charge(db_id, cid, month, year, name, qty, price)
            self._ppe_row_ids[row] = new_id
        if errors:
            QMessageBox.warning(self, "Validation", "\n".join(errors))
        else:
            QMessageBox.information(self, "Saved", "PPE charges saved.")
        self._load_charges()

    # ── Diesel helpers ────────────────────────────────────────────────────────

    def _add_diesel_row(self):
        if not self.contractor_combo.currentData():
            return
        self.diesel_table.blockSignals(True)
        self._append_charge_row(self.diesel_table, self._diesel_row_ids, None, "", 0, 0)
        self.diesel_table.blockSignals(False)

    def _delete_diesel_row(self):
        row = self.diesel_table.currentRow()
        if row < 0:
            return
        db_id = self._diesel_row_ids[row]
        if db_id:
            reply = QMessageBox.question(self, "Confirm", "Delete this row?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply != QMessageBox.StandardButton.Yes:
                return
            m.delete_diesel_charge(db_id)
        self.diesel_table.removeRow(row)
        self._diesel_row_ids.pop(row)
        self._recalc_table(self.diesel_table, self.diesel_total_lbl)

    def _save_diesel(self):
        cid = self.contractor_combo.currentData()
        month = self.month_combo.currentData()
        year = self.year_combo.currentData()
        if not cid:
            return
        errors = []
        for row in range(self.diesel_table.rowCount()):
            desc = (self.diesel_table.item(row, 0) or QTableWidgetItem("")).text().strip()
            if not desc:
                errors.append(f"Row {row + 1}: Description is empty.")
                continue
            try:
                qty = float((self.diesel_table.item(row, 1) or QTableWidgetItem("0")).text())
                price = float((self.diesel_table.item(row, 2) or QTableWidgetItem("0")).text())
            except ValueError:
                errors.append(f"Row {row + 1}: Invalid number.")
                continue
            db_id = self._diesel_row_ids[row]
            new_id = m.upsert_diesel_charge(db_id, cid, month, year, desc, qty, price)
            self._diesel_row_ids[row] = new_id
        if errors:
            QMessageBox.warning(self, "Validation", "\n".join(errors))
        else:
            QMessageBox.information(self, "Saved", "Diesel charges saved.")
        self._load_charges()
