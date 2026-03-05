import os
from datetime import date

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QComboBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QFileDialog, QMessageBox, QFrame, QLineEdit
)
from PyQt6.QtCore import Qt

import db.models as m
from export.excel_exporter import generate_invoice
from .styles import (PAGE_TITLE_STYLE, TABLE_STYLE, BTN_PRIMARY,
                     BTN_SUCCESS, COMBO_STYLE, LABEL_MUTED)

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]


class InvoicePreviewView(QWidget):
    def __init__(self):
        super().__init__()
        self._project_id = None
        self._project_name = ""
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(16)

        title = QLabel("Invoice / Export")
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
        self.contractor_combo.setMinimumWidth(220)
        self.contractor_combo.addItem("All Contractors", -1)
        filter_bar.addWidget(self.contractor_combo)

        filter_bar.addWidget(QLabel("Month:"))
        self.month_combo = QComboBox()
        self.month_combo.setStyleSheet(COMBO_STYLE)
        for i, name in enumerate(MONTHS):
            self.month_combo.addItem(name, i + 1)
        current = date.today().month
        self.month_combo.setCurrentIndex(current - 1)
        filter_bar.addWidget(self.month_combo)

        filter_bar.addWidget(QLabel("Year:"))
        self.year_combo = QComboBox()
        self.year_combo.setStyleSheet(COMBO_STYLE)
        current_year = date.today().year
        for y in range(current_year - 2, current_year + 3):
            self.year_combo.addItem(str(y), y)
        self.year_combo.setCurrentText(str(current_year))
        filter_bar.addWidget(self.year_combo)

        # ── Exchange rate ─────────────────────────────────────────────────
        filter_bar.addSpacing(20)
        rate_lbl = QLabel("1 USD =")
        rate_lbl.setStyleSheet("font-weight: bold;")
        filter_bar.addWidget(rate_lbl)

        self.rate_input = QLineEdit()
        saved_rate = m.get_setting("usd_tl_rate", "1.00")
        self.rate_input.setText(saved_rate)
        self.rate_input.setMaximumWidth(80)
        self.rate_input.setToolTip(
            "TL per USD exchange rate.\n"
            "Deductions entered in TL will be divided by this rate to get USD equivalent."
        )
        self.rate_input.textChanged.connect(self._save_rate)
        filter_bar.addWidget(self.rate_input)

        tl_lbl = QLabel("TL")
        tl_lbl.setStyleSheet("font-weight: bold; color: #b71c1c;")
        filter_bar.addWidget(tl_lbl)

        preview_btn = QPushButton("Preview")
        preview_btn.setStyleSheet(BTN_PRIMARY)
        preview_btn.clicked.connect(self._preview)
        filter_bar.addSpacing(12)
        filter_bar.addWidget(preview_btn)
        filter_bar.addStretch()
        layout.addLayout(filter_bar)

        # ── Summary table ────────────────────────────────────────────────
        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE)
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "Contractor", "Type", "Work Total (USD)",
            "Deductions (TL)", "Deductions (USD)", "Net Payable (USD)"
        ])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for col, w in [(1, 100), (2, 140), (3, 140), (4, 140), (5, 150)]:
            self.table.setColumnWidth(col, w)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        # ── Grand total footer ───────────────────────────────────────────
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #cfd8dc;")
        layout.addWidget(sep)

        footer_row = QHBoxLayout()
        self.grand_work_lbl    = QLabel("Work Total: $0.00")
        self.grand_deduct_tl_lbl = QLabel("Deductions: ₺0.00")
        self.grand_deduct_usd_lbl = QLabel("Deductions (USD): $0.00")
        self.grand_net_lbl     = QLabel("Net Payable (USD): $0.00")
        self.grand_net_lbl.setStyleSheet("font-weight: bold; font-size: 15px; color: #1b5e20;")
        self.grand_net_tl_lbl  = QLabel("Net Payable (TL): ₺0.00")
        self.grand_net_tl_lbl.setStyleSheet("font-weight: bold; font-size: 13px; color: #b71c1c;")

        for lbl in [self.grand_work_lbl, self.grand_deduct_tl_lbl,
                    self.grand_deduct_usd_lbl]:
            footer_row.addWidget(lbl)
        footer_row.addStretch()
        footer_row.addWidget(self.grand_net_tl_lbl)
        footer_row.addSpacing(20)
        footer_row.addWidget(self.grand_net_lbl)
        layout.addLayout(footer_row)

        # ── Export button ────────────────────────────────────────────────
        export_row = QHBoxLayout()
        self.export_btn = QPushButton("Export to Excel (.xlsx)")
        self.export_btn.setStyleSheet(BTN_SUCCESS)
        self.export_btn.setMinimumHeight(40)
        self.export_btn.clicked.connect(self._export)
        export_row.addStretch()
        export_row.addWidget(self.export_btn)
        layout.addLayout(export_row)

    # ── Public API ────────────────────────────────────────────────────────────

    def set_project(self, project_id: int, project_name: str):
        self._project_id = project_id
        self._project_name = project_name
        self.ctx_label.setText(f"Project: {project_name}")
        self._refresh_contractors()

    # ── Internal helpers ──────────────────────────────────────────────────────

    def _save_rate(self, text: str):
        try:
            float(text)
            m.set_setting("usd_tl_rate", text.strip())
        except ValueError:
            pass

    def _get_rate(self) -> float:
        try:
            return float(self.rate_input.text())
        except ValueError:
            return 1.0

    def _refresh_contractors(self):
        self.contractor_combo.blockSignals(True)
        self.contractor_combo.clear()
        self.contractor_combo.addItem("All Contractors", -1)
        if self._project_id:
            for c in m.get_contractors(self._project_id):
                self.contractor_combo.addItem(
                    f"{c['name']} ({c['type'].capitalize()})", c["id"]
                )
        self.contractor_combo.blockSignals(False)

    def _preview(self):
        if not self._project_id:
            QMessageBox.information(self, "Select Project", "Please select a project first.")
            return

        selected_cid = self.contractor_combo.currentData()
        month = self.month_combo.currentData()
        year  = self.year_combo.currentData()
        rate  = self._get_rate()

        contractors = m.get_contractors(self._project_id)
        if selected_cid != -1:
            contractors = [c for c in contractors if c["id"] == selected_cid]

        rows = []
        for contractor in contractors:
            cid    = contractor["id"]
            ctype  = contractor["type"]
            rate_m = contractor["rate_per_meter"]
            rate_s = contractor["standby_hour_rate"]

            entries  = m.get_drilling_entries(cid, month, year)
            standby  = m.get_standby_entries(cid, month, year)
            work_usd = (
                sum(e["meters_drilled"] * rate_m for e in entries) +
                sum(e["hours"] * rate_s for e in standby)
            )

            if ctype == "underground":
                charges = m.get_ppe_charges(cid, month, year)
            else:
                charges = m.get_diesel_charges(cid, month, year)
            deduct_tl  = sum(r["quantity"] * r["unit_price"] for r in charges)
            deduct_usd = deduct_tl / rate if rate else 0.0

            rows.append({
                "name":       contractor["name"],
                "type":       ctype,
                "work":       work_usd,
                "deduct_tl":  deduct_tl,
                "deduct_usd": deduct_usd,
                "net":        work_usd - deduct_usd,
            })

        self.table.setRowCount(len(rows))
        gw = gdt = gdu = gn = 0.0
        for r, data in enumerate(rows):
            self.table.setItem(r, 0, QTableWidgetItem(data["name"]))
            self.table.setItem(r, 1, QTableWidgetItem(data["type"].capitalize()))
            self.table.setItem(r, 2, QTableWidgetItem(f"${data['work']:,.2f}"))
            self.table.setItem(r, 3, QTableWidgetItem(f"₺{data['deduct_tl']:,.2f}"))
            self.table.setItem(r, 4, QTableWidgetItem(f"${data['deduct_usd']:,.2f}"))
            net_item = QTableWidgetItem(f"${data['net']:,.2f}")
            net_item.setForeground(Qt.GlobalColor.darkGreen)
            self.table.setItem(r, 5, net_item)
            gw  += data["work"]
            gdt += data["deduct_tl"]
            gdu += data["deduct_usd"]
            gn  += data["net"]

        self.grand_work_lbl.setText(f"Work Total: ${gw:,.2f}")
        self.grand_deduct_tl_lbl.setText(f"Deductions: ₺{gdt:,.2f}")
        self.grand_deduct_usd_lbl.setText(f"Deductions (USD): ${gdu:,.2f}")
        self.grand_net_lbl.setText(f"Net Payable (USD): ${gn:,.2f}")
        self.grand_net_tl_lbl.setText(f"Net Payable (TL): ₺{gn * rate:,.2f}")

    def _export(self):
        if not self._project_id:
            QMessageBox.information(self, "Select Project", "Please select a project first.")
            return

        month = self.month_combo.currentData()
        year  = self.year_combo.currentData()
        selected_cid = self.contractor_combo.currentData()
        rate = self._get_rate()
        period_str = f"{MONTHS[month - 1]}_{year}"
        default_name = f"Invoice_{self._project_name}_{period_str}.xlsx".replace(" ", "_")

        path, _ = QFileDialog.getSaveFileName(
            self, "Save Invoice", default_name,
            "Excel Files (*.xlsx)"
        )
        if not path:
            return
        if not path.endswith(".xlsx"):
            path += ".xlsx"

        try:
            generate_invoice(self._project_id, selected_cid, month, year, path,
                             usd_tl_rate=rate)
            QMessageBox.information(
                self, "Export Complete",
                f"Invoice exported successfully:\n{path}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Export Failed", str(e))
