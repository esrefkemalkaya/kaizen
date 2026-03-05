from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QPushButton, QStackedWidget, QLabel, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont


NAV_ITEMS = [
    ("Projects",         0),
    ("Contractors",      1),
    ("Drilling Entries", 2),
    ("Charges",          3),
    ("Invoice / Export", 4),
]

SIDEBAR_STYLE = """
    QWidget#sidebar {
        background-color: #1e2a38;
    }
"""

NAV_BTN_STYLE = """
    QPushButton {
        color: #b0bec5;
        background: transparent;
        border: none;
        text-align: left;
        padding: 12px 20px;
        font-size: 14px;
    }
    QPushButton:hover {
        background-color: #263547;
        color: #ffffff;
    }
    QPushButton[active="true"] {
        background-color: #2e7d32;
        color: #ffffff;
        font-weight: bold;
    }
"""

HEADER_STYLE = """
    QLabel {
        color: #ffffff;
        font-size: 18px;
        font-weight: bold;
        padding: 18px 20px 10px 20px;
    }
"""

DIVIDER_STYLE = "background-color: #37474f;"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Kaizen — Drilling Invoice Manager")
        self.setMinimumSize(1100, 700)
        self._nav_buttons = []
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Sidebar ──────────────────────────────────────────────────────
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(210)
        sidebar.setStyleSheet(SIDEBAR_STYLE)
        sb_layout = QVBoxLayout(sidebar)
        sb_layout.setContentsMargins(0, 0, 0, 0)
        sb_layout.setSpacing(0)

        title = QLabel("⛏  Kaizen")
        title.setStyleSheet(HEADER_STYLE)
        sb_layout.addWidget(title)

        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setFixedHeight(1)
        divider.setStyleSheet(DIVIDER_STYLE)
        sb_layout.addWidget(divider)
        sb_layout.addSpacing(8)

        for label, idx in NAV_ITEMS:
            btn = QPushButton(label)
            btn.setStyleSheet(NAV_BTN_STYLE)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(lambda checked, i=idx: self._navigate(i))
            self._nav_buttons.append(btn)
            sb_layout.addWidget(btn)

        sb_layout.addStretch()
        root.addWidget(sidebar)

        # ── Page stack ────────────────────────────────────────────────────
        self.stack = QStackedWidget()
        self._init_pages()
        root.addWidget(self.stack)

        self._navigate(0)

    def _init_pages(self):
        from .projects_view import ProjectsView
        from .contractors_view import ContractorsView
        from .drilling_entry_view import DrillingEntryView
        from .charges_view import ChargesView
        from .invoice_preview import InvoicePreviewView

        self.projects_view = ProjectsView()
        self.contractors_view = ContractorsView()
        self.drilling_view = DrillingEntryView()
        self.charges_view = ChargesView()
        self.invoice_view = InvoicePreviewView()

        for view in [self.projects_view, self.contractors_view,
                     self.drilling_view, self.charges_view, self.invoice_view]:
            self.stack.addWidget(view)

        # Wire project selection → contractors view context
        self.projects_view.project_selected.connect(self.contractors_view.set_project)
        self.projects_view.project_selected.connect(self.drilling_view.set_project)
        self.projects_view.project_selected.connect(self.charges_view.set_project)
        self.projects_view.project_selected.connect(self.invoice_view.set_project)

    def _navigate(self, index: int):
        self.stack.setCurrentIndex(index)
        for i, btn in enumerate(self._nav_buttons):
            btn.setProperty("active", i == index)
            btn.style().unpolish(btn)
            btn.style().polish(btn)
