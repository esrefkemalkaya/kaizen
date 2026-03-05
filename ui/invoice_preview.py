import os
from datetime import date

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QComboBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QFileDialog, QMessageBox, QFrame
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

        preview_btn = QPushButton("Preview")
        preview_btn.setStyleSheet(BTN_PRIMARY)
        preview_btn.clicked.connect(self._preview)
        filter_bar.addWidget(preview_btn)
        filter_bar.addStretch()
        layout.addLayout(filter_bar)

        # ── Summary table ────────────────────────────────────────────────
        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels([
            "Contractor", "Type", "Drilling Total", "Deductions", "Net Payable"
        ])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(1, 110)
        self.table.setColumnWidth(2, 140)
        self.table.setColumnWidth(3, 130)
        self.table.setColumnWidth(4, 140)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        # ── Grand total footer ───────────────────────────────────────────
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #cfd8dc;")
        layout.addWidget(sep)

        footer_row = QHBoxLayout()
        self.grand_drill_lbl  = QLabel("Drilling Total: $0.00")
        self.grand_deduct_lbl = QLabel("Total Deductions: $0.00")
        self.grand_net_lbl    = QLabel("Net Payable: $0.00")
        self.grand_net_lbl.setStyleSheet("font-weight: bold; font-size: 15px; color: #1b5e20;")
        footer_row.addWidget(self.grand_drill_lbl)
        footer_row.addWidget(self.grand_deduct_lbl)
        footer_row.addStretch()
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

        contractors = m.get_contractors(self._project_id)
        if selected_cid != -1:
            contractors = [c for c in contractors if c["id"] == selected_cid]

        rows = []
        for contractor in contractors:
            cid    = contractor["id"]
            ctype  = contractor["type"]
            rate_m = contractor["rate_per_meter"]
            rate_s = contractor["standby_hour_rate"]

            entries = m.get_drilling_entries(cid, month, year)
            drill_total = sum(
                e["meters_drilled"] * rate_m + e["standby_hours"] * rate_s
                for e in entries
            )

            if ctype == "underground":
                charges = m.get_ppe_charges(cid, month, year)
            else:
                charges = m.get_diesel_charges(cid, month, year)
            deduct_total = sum(r["quantity"] * r["unit_price"] for r in charges)

            rows.append({
                "name":       contractor["name"],
                "type":       ctype,
                "drilling":   drill_total,
                "deductions": deduct_total,
                "net":        drill_total - deduct_total,
            })

        self.table.setRowCount(len(rows))
        grand_d = grand_ded = grand_net = 0.0
        for r, data in enumerate(rows):
            self.table.setItem(r, 0, QTableWidgetItem(data["name"]))
            self.table.setItem(r, 1, QTableWidgetItem(data["type"].capitalize()))
            self.table.setItem(r, 2, QTableWidgetItem(f"${data['drilling']:,.2f}"))
            self.table.setItem(r, 3, QTableWidgetItem(f"${data['deductions']:,.2f}"))
            net_item = QTableWidgetItem(f"${data['net']:,.2f}")
            net_item.setForeground(Qt.GlobalColor.darkGreen)
            self.table.setItem(r, 4, net_item)
            grand_d   += data["drilling"]
            grand_ded += data["deductions"]
            grand_net += data["net"]

        self.grand_drill_lbl.setText(f"Drilling Total: ${grand_d:,.2f}")
        self.grand_deduct_lbl.setText(f"Total Deductions: ${grand_ded:,.2f}")
        self.grand_net_lbl.setText(f"Net Payable: ${grand_net:,.2f}")

    def _export(self):
        if not self._project_id:
            QMessageBox.information(self, "Select Project", "Please select a project first.")
            return

        month = self.month_combo.currentData()
        year  = self.year_combo.currentData()
        selected_cid = self.contractor_combo.currentData()
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
            generate_invoice(self._project_id, selected_cid, month, year, path)
            QMessageBox.information(
                self, "Export Complete",
                f"Invoice exported successfully:\n{path}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Export Failed", str(e))
