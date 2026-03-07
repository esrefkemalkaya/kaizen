from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QDialog, QFormLayout, QLineEdit, QComboBox,
    QDialogButtonBox, QMessageBox, QLabel, QDoubleSpinBox,
    QGroupBox, QGridLayout, QSpinBox
)
from PyQt6.QtCore import Qt
from datetime import date as _date

import db.models as m
from .styles import PAGE_TITLE_STYLE, TABLE_STYLE, BTN_PRIMARY, BTN_DANGER, BTN_SUCCESS, LABEL_MUTED, COMBO_STYLE

MONTHS_TR = [
    "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
    "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"
]
MONTHS_EN = ["January", "February", "March", "April", "May", "June",
             "July", "August", "September", "October", "November", "December"]

MACHINES = ["GEO 900E-1", "GEO 900E-2", "GEO 900E-3", "GEO 900E-5"]


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
        layout.setSpacing(14)

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

        # ── Period Parameters ──────────────────────────────────────────────
        period_box = QGroupBox("Dönem Parametreleri — Period Parameters")
        period_box.setStyleSheet("""
            QGroupBox { font-weight: bold; font-size: 13px; color: #1565c0;
                        border: 1.5px solid #90caf9; border-radius: 6px;
                        margin-top: 8px; padding-top: 10px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
        """)
        pb_layout = QVBoxLayout(period_box)
        pb_layout.setSpacing(10)

        # Contractor + period selector row
        sel_row = QHBoxLayout()
        sel_row.addWidget(QLabel("Contractor:"))
        self.ps_contractor_combo = QComboBox()
        self.ps_contractor_combo.setStyleSheet(COMBO_STYLE)
        self.ps_contractor_combo.setMinimumWidth(200)
        sel_row.addWidget(self.ps_contractor_combo)
        sel_row.addSpacing(12)
        sel_row.addWidget(QLabel("Month:"))
        self.ps_month_combo = QComboBox()
        self.ps_month_combo.setStyleSheet(COMBO_STYLE)
        for i, name in enumerate(MONTHS_EN):
            self.ps_month_combo.addItem(name, i + 1)
        self.ps_month_combo.setCurrentIndex(_date.today().month - 1)
        sel_row.addWidget(self.ps_month_combo)
        sel_row.addSpacing(6)
        sel_row.addWidget(QLabel("Year:"))
        self.ps_year_spin = QSpinBox()
        self.ps_year_spin.setRange(2020, 2035)
        self.ps_year_spin.setValue(_date.today().year)
        self.ps_year_spin.setFixedWidth(70)
        sel_row.addWidget(self.ps_year_spin)
        load_ps_btn = QPushButton("Load")
        load_ps_btn.setStyleSheet(BTN_PRIMARY)
        load_ps_btn.clicked.connect(self._ps_load)
        sel_row.addWidget(load_ps_btn)
        sel_row.addStretch()
        pb_layout.addLayout(sel_row)

        # Parameter grid
        grid = QGridLayout()
        grid.setHorizontalSpacing(20)
        grid.setVerticalSpacing(8)

        grid.addWidget(QLabel("Dönem Adı (H2):"), 0, 0)
        self.ps_donem = QLineEdit()
        self.ps_donem.setPlaceholderText("örn: ŞUBAT 2026")
        self.ps_donem.setMinimumWidth(150)
        grid.addWidget(self.ps_donem, 0, 1)

        grid.addWidget(QLabel("Birim Fiyat (H26) USD/m:"), 0, 2)
        self.ps_rate_m = QDoubleSpinBox()
        self.ps_rate_m.setRange(0, 9999)
        self.ps_rate_m.setDecimals(2)
        self.ps_rate_m.setPrefix("$ ")
        grid.addWidget(self.ps_rate_m, 0, 3)

        grid.addWidget(QLabel("USD/TL Kuru (H53):"), 1, 0)
        self.ps_kur = QDoubleSpinBox()
        self.ps_kur.setRange(0, 9999)
        self.ps_kur.setDecimals(4)
        self.ps_kur.setValue(43.49)
        grid.addWidget(self.ps_kur, 1, 1)

        grid.addWidget(QLabel("Bekleme Birim Fiyat (G45-48) USD/sa:"), 1, 2)
        self.ps_sb_rate = QDoubleSpinBox()
        self.ps_sb_rate.setRange(0, 9999)
        self.ps_sb_rate.setDecimals(2)
        self.ps_sb_rate.setValue(75.0)
        self.ps_sb_rate.setPrefix("$ ")
        grid.addWidget(self.ps_sb_rate, 1, 3)

        grid.addWidget(QLabel("Kuyuda Kalan (H41) USD:"), 2, 0)
        self.ps_kuyuda_kalan = QDoubleSpinBox()
        self.ps_kuyuda_kalan.setRange(-999999, 999999)
        self.ps_kuyuda_kalan.setDecimals(2)
        self.ps_kuyuda_kalan.setPrefix("$ ")
        grid.addWidget(self.ps_kuyuda_kalan, 2, 1)

        pb_layout.addLayout(grid)

        # Target meters per machine
        target_row = QHBoxLayout()
        target_row.addWidget(QLabel("Hedef Metre / makine (C33:C36):"))
        self.ps_targets: dict[str, QDoubleSpinBox] = {}
        for machine in MACHINES:
            target_row.addSpacing(8)
            target_row.addWidget(QLabel(machine + ":"))
            spin = QDoubleSpinBox()
            spin.setRange(0, 9999)
            spin.setDecimals(0)
            spin.setValue(700)
            spin.setFixedWidth(80)
            self.ps_targets[machine] = spin
            target_row.addWidget(spin)
        target_row.addStretch()
        pb_layout.addLayout(target_row)

        save_ps_btn = QPushButton("Kaydet — Save Period Parameters")
        save_ps_btn.setStyleSheet(BTN_SUCCESS)
        save_ps_btn.clicked.connect(self._ps_save)
        pb_layout.addWidget(save_ps_btn)

        layout.addWidget(period_box)

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
        # Populate period params contractor combo
        self.ps_contractor_combo.blockSignals(True)
        self.ps_contractor_combo.clear()
        for c in rows:
            self.ps_contractor_combo.addItem(c["name"], c["id"])
        self.ps_contractor_combo.blockSignals(False)

    # ── Period Parameters ──────────────────────────────────────────────────

    def _ps_load(self):
        cid = self.ps_contractor_combo.currentData()
        if not cid:
            return
        month = self.ps_month_combo.currentData()
        year  = self.ps_year_spin.value()
        ps = m.get_period_settings(cid, month, year)
        # Populate rate_per_meter from contractor table
        contractor = m.get_contractor(cid)
        if contractor:
            self.ps_rate_m.setValue(contractor["rate_per_meter"])
        self.ps_donem.setText(ps.get("donem_adi", ""))
        self.ps_kur.setValue(float(ps.get("exchange_rate", 43.49) or 43.49))
        self.ps_kuyuda_kalan.setValue(float(ps.get("kuyuda_kalan", 0) or 0))
        self.ps_sb_rate.setValue(float(ps.get("standby_rate", 75) or 75))
        targets_map = {
            "GEO 900E-1": "target_geo1",
            "GEO 900E-2": "target_geo2",
            "GEO 900E-3": "target_geo3",
            "GEO 900E-5": "target_geo5",
        }
        for machine, key in targets_map.items():
            if machine in self.ps_targets:
                self.ps_targets[machine].setValue(float(ps.get(key, 700) or 700))

    def _ps_save(self):
        cid = self.ps_contractor_combo.currentData()
        if not cid:
            QMessageBox.information(self, "Select", "Select a contractor first.")
            return
        month = self.ps_month_combo.currentData()
        year  = self.ps_year_spin.value()
        # Also update rate_per_meter on contractor if changed
        contractor = m.get_contractor(cid)
        if contractor:
            new_rate = self.ps_rate_m.value()
            if abs(new_rate - contractor["rate_per_meter"]) > 0.001:
                m.update_contractor(cid, contractor["name"], contractor["type"],
                                    new_rate, contractor["standby_hour_rate"])
        m.upsert_period_settings(
            cid, month, year,
            donem_adi=self.ps_donem.text().strip(),
            exchange_rate=self.ps_kur.value(),
            kuyuda_kalan=self.ps_kuyuda_kalan.value(),
            target_geo1=self.ps_targets["GEO 900E-1"].value(),
            target_geo2=self.ps_targets["GEO 900E-2"].value(),
            target_geo3=self.ps_targets["GEO 900E-3"].value(),
            target_geo5=self.ps_targets["GEO 900E-5"].value(),
            standby_rate=self.ps_sb_rate.value(),
        )
        QMessageBox.information(self, "Saved",
            f"Period parameters saved for {self.ps_month_combo.currentText()} {year}.")

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
