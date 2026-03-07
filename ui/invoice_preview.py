"""
Invoice / Export view — Giriş Şablonu based export.

Workflow:
  1. Set template path (saved in DB settings) once.
  2. Select contractor + month/year.
  3. Click "Preview" to see a quick summary.
  4. Click "Generate Giriş Şablonu" → choose save path → app fills the
     template and opens it (or just saves it, user opens manually).

Period parameters (target metres, exchange rate, etc.) are set in the
Contractors page → "Dönem Parametreleri" section.
"""

import os
import subprocess
import sys
from datetime import date

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QComboBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QFileDialog, QMessageBox, QFrame,
    QLineEdit, QGroupBox
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

import db.models as m
from .styles import (PAGE_TITLE_STYLE, TABLE_STYLE, BTN_PRIMARY,
                     BTN_SUCCESS, BTN_DANGER, COMBO_STYLE, LABEL_MUTED)

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

MACHINES = ["GEO 900E-1", "GEO 900E-2", "GEO 900E-3", "GEO 900E-5"]


class InvoicePreviewView(QWidget):
    def __init__(self):
        super().__init__()
        self._project_id = None
        self._project_name = ""
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(14)

        title = QLabel("Export — Giriş Şablonu")
        title.setStyleSheet(PAGE_TITLE_STYLE)
        layout.addWidget(title)

        self.ctx_label = QLabel("No project selected.")
        self.ctx_label.setStyleSheet(LABEL_MUTED)
        layout.addWidget(self.ctx_label)

        # ── Template path ─────────────────────────────────────────────────
        tpl_box = QGroupBox("Excel Şablon Dosyası — Template File (.xlsx)")
        tpl_box.setStyleSheet("""
            QGroupBox { font-weight: bold; font-size: 12px; color: #37474f;
                        border: 1px solid #b0bec5; border-radius: 5px;
                        margin-top: 6px; padding-top: 8px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; }
        """)
        tpl_layout = QHBoxLayout(tpl_box)
        tpl_layout.setSpacing(8)

        self.tpl_path_edit = QLineEdit()
        self.tpl_path_edit.setPlaceholderText("Giriş/Çıktı şablon dosyasının yolunu seçin…")
        self.tpl_path_edit.setText(m.get_setting("giris_template_path", ""))
        self.tpl_path_edit.setReadOnly(True)
        tpl_layout.addWidget(self.tpl_path_edit)

        browse_btn = QPushButton("Şablon Seç…")
        browse_btn.setStyleSheet(BTN_PRIMARY)
        browse_btn.clicked.connect(self._browse_template)
        tpl_layout.addWidget(browse_btn)

        tpl_hint = QLabel("(Formüller hazır şablon. Uygulama sadece giriş hücrelerini yazar.)")
        tpl_hint.setStyleSheet("color: #78909c; font-size: 11px; font-weight: normal;")
        tpl_layout.addWidget(tpl_hint)

        layout.addWidget(tpl_box)

        # ── Filter bar ───────────────────────────────────────────────────
        filter_bar = QHBoxLayout()
        filter_bar.setSpacing(12)

        filter_bar.addWidget(QLabel("Contractor:"))
        self.contractor_combo = QComboBox()
        self.contractor_combo.setStyleSheet(COMBO_STYLE)
        self.contractor_combo.setMinimumWidth(220)
        filter_bar.addWidget(self.contractor_combo)

        filter_bar.addWidget(QLabel("Month:"))
        self.month_combo = QComboBox()
        self.month_combo.setStyleSheet(COMBO_STYLE)
        for i, name in enumerate(MONTHS):
            self.month_combo.addItem(name, i + 1)
        self.month_combo.setCurrentIndex(date.today().month - 1)
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
        filter_bar.addSpacing(12)
        filter_bar.addWidget(preview_btn)
        filter_bar.addStretch()
        layout.addLayout(filter_bar)

        # ── Summary table ────────────────────────────────────────────────
        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE)
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Makine", "Kuyu Sayısı", "Toplam İlerleme (m)",
            "Hedef (m)", "Oran %", "Net Bekleme (h)", "Net Ödeme (USD)"
        ])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for col, w in [(1, 90), (2, 140), (3, 90), (4, 70), (5, 120), (6, 140)]:
            self.table.setColumnWidth(col, w)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        # ── Summary footer ───────────────────────────────────────────────
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #cfd8dc;")
        layout.addWidget(sep)

        self.summary_lbl = QLabel("")
        self.summary_lbl.setStyleSheet("font-size: 12px; color: #37474f;")
        layout.addWidget(self.summary_lbl)

        self.period_warn_lbl = QLabel("")
        self.period_warn_lbl.setStyleSheet("color: #e65100; font-size: 11px;")
        layout.addWidget(self.period_warn_lbl)

        # ── Export button ────────────────────────────────────────────────
        export_row = QHBoxLayout()
        self.export_btn = QPushButton("▶  Generate Giriş Şablonu (.xlsx)")
        self.export_btn.setStyleSheet(BTN_SUCCESS)
        self.export_btn.setMinimumHeight(44)
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

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _browse_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Giriş Template", "",
            "Excel Files (*.xlsx *.xlsm);;All Files (*)"
        )
        if path:
            self.tpl_path_edit.setText(path)
            m.set_setting("giris_template_path", path)

    def _refresh_contractors(self):
        self.contractor_combo.blockSignals(True)
        self.contractor_combo.clear()
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
        cid   = self.contractor_combo.currentData()
        month = self.month_combo.currentData()
        year  = self.year_combo.currentData()
        if not cid:
            return

        ps        = m.get_period_settings(cid, month, year)
        entries   = m.get_drilling_entries(cid, month, year)
        sb_hours  = m.get_standby_net_hours_per_rig(cid, month, year)
        contractor = m.get_contractor(cid)
        rate_m    = contractor["rate_per_meter"] if contractor else 0.0
        rate_s    = ps.get("standby_rate", 75.0) or 75.0

        target_map = {
            "GEO 900E-1": ps.get("target_geo1", 700.0) or 700.0,
            "GEO 900E-2": ps.get("target_geo2", 700.0) or 700.0,
            "GEO 900E-3": ps.get("target_geo3", 700.0) or 700.0,
            "GEO 900E-5": ps.get("target_geo5", 700.0) or 700.0,
        }

        # Per-machine actual metres
        from collections import defaultdict
        rig_meters: dict[str, float] = defaultdict(float)
        rig_holes:  dict[str, int]   = defaultdict(int)
        for e in entries:
            rig = (e["rig_name"] or "").strip()
            rig_meters[rig] += e["meters_drilled"]
            rig_holes[rig]  += 1

        rows = []
        total_net_usd = 0.0
        for rig in MACHINES:
            actual  = rig_meters.get(rig, 0.0)
            target  = target_map[rig]
            holes   = rig_holes.get(rig, 0)
            sb_h    = sb_hours.get(rig, 0.0)
            ratio   = actual / target if target else 0.0

            # Bonus/ceza logic
            if actual >= target:
                net_usd = target * rate_m + (actual - target) * rate_m * 1.2
            elif actual >= target * 0.5:
                net_usd = actual * rate_m - (target - actual) * rate_m * 0.2
            else:
                net_usd = actual * rate_m * 0.8

            net_usd += sb_h * rate_s
            total_net_usd += net_usd
            rows.append((rig, holes, actual, target, ratio, sb_h, net_usd))

        self.table.setRowCount(len(rows))
        for r, (rig, holes, actual, target, ratio, sb_h, net_usd) in enumerate(rows):
            self.table.setItem(r, 0, QTableWidgetItem(rig))
            self.table.setItem(r, 1, QTableWidgetItem(str(holes)))
            self.table.setItem(r, 2, QTableWidgetItem(f"{actual:,.2f}"))
            self.table.setItem(r, 3, QTableWidgetItem(f"{target:,.0f}"))

            pct = QTableWidgetItem(f"{ratio*100:.0f}%")
            if ratio >= 1.0:
                pct.setForeground(QColor("#1b5e20"))
            elif ratio >= 0.5:
                pct.setForeground(QColor("#e65100"))
            else:
                pct.setForeground(QColor("#b71c1c"))
            self.table.setItem(r, 4, pct)

            self.table.setItem(r, 5, QTableWidgetItem(f"{sb_h:,.2f}"))
            net_item = QTableWidgetItem(f"${net_usd:,.2f}")
            net_item.setForeground(QColor("#1b5e20"))
            self.table.setItem(r, 6, net_item)

        donem = ps.get("donem_adi", "") or f"{MONTHS[month-1]} {year}"
        kur   = ps.get("exchange_rate", 0.0) or 0.0
        self.summary_lbl.setText(
            f"Dönem: {donem}  |  Toplam Net: ${total_net_usd:,.2f}"
            + (f"  =  ₺{total_net_usd * kur:,.2f}" if kur else "")
        )

        # Warn if period params missing
        warnings = []
        if not ps.get("donem_adi"):
            warnings.append("Dönem adı girilmemiş")
        if not ps.get("exchange_rate"):
            warnings.append("USD/TL kuru girilmemiş")
        if not self.tpl_path_edit.text():
            warnings.append("Şablon dosyası seçilmemiş")
        self.period_warn_lbl.setText(
            "⚠  " + "  |  ".join(warnings) if warnings else ""
        )

    def _export(self):
        if not self._project_id:
            QMessageBox.information(self, "Select Project", "Please select a project first.")
            return

        tpl_path = self.tpl_path_edit.text().strip()
        if not tpl_path:
            QMessageBox.warning(self, "No Template",
                "Şablon dosyası seçilmemiş.\n"
                "Üstteki 'Şablon Seç…' butonundan Giriş/Çıktı şablonunu seçin.")
            return

        cid   = self.contractor_combo.currentData()
        month = self.month_combo.currentData()
        year  = self.year_combo.currentData()
        if not cid:
            QMessageBox.information(self, "Select", "Select a contractor first.")
            return

        ps = m.get_period_settings(cid, month, year)
        period_str = (ps.get("donem_adi") or f"{MONTHS[month-1]}_{year}").replace(" ", "_")
        contractor = m.get_contractor(cid)
        cname = contractor["name"].replace(" ", "_") if contractor else "contractor"
        default_name = f"Giris_{cname}_{period_str}.xlsx"

        out_path, _ = QFileDialog.getSaveFileName(
            self, "Çıktı Dosyasını Kaydet", default_name,
            "Excel Files (*.xlsx)"
        )
        if not out_path:
            return
        if not out_path.endswith(".xlsx"):
            out_path += ".xlsx"

        try:
            from export.giris_exporter import fill_template
            fill_template(tpl_path, out_path, cid, month, year)

            reply = QMessageBox.information(
                self, "Tamamlandı",
                f"Şablon dolduruldu:\n{out_path}\n\n"
                "Excel'de açmak ister misiniz?",
                QMessageBox.StandardButton.Open | QMessageBox.StandardButton.Ok,
                QMessageBox.StandardButton.Ok
            )
            if reply == QMessageBox.StandardButton.Open:
                _open_file(out_path)
        except Exception as e:
            QMessageBox.critical(self, "Export Failed", str(e))


def _open_file(path: str):
    """Open file with default OS application."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass
