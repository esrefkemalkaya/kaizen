PAGE_TITLE_STYLE = """
    font-size: 22px;
    font-weight: bold;
    color: #1e2a38;
    margin-bottom: 4px;
"""

TABLE_STYLE = """
    QTableWidget {
        border: 1px solid #cfd8dc;
        gridline-color: #eceff1;
        font-size: 13px;
    }
    QHeaderView::section {
        background-color: #37474f;
        color: white;
        padding: 6px;
        font-weight: bold;
        border: none;
    }
    QTableWidget::item:selected {
        background-color: #1565c0;
        color: white;
    }
"""

BTN_PRIMARY = """
    QPushButton {
        background-color: #1565c0;
        color: white;
        border: none;
        padding: 7px 18px;
        border-radius: 4px;
        font-size: 13px;
    }
    QPushButton:hover { background-color: #1976d2; }
    QPushButton:pressed { background-color: #0d47a1; }
"""

BTN_DANGER = """
    QPushButton {
        background-color: #c62828;
        color: white;
        border: none;
        padding: 7px 18px;
        border-radius: 4px;
        font-size: 13px;
    }
    QPushButton:hover { background-color: #d32f2f; }
    QPushButton:pressed { background-color: #b71c1c; }
"""

BTN_SUCCESS = """
    QPushButton {
        background-color: #2e7d32;
        color: white;
        border: none;
        padding: 7px 18px;
        border-radius: 4px;
        font-size: 13px;
    }
    QPushButton:hover { background-color: #388e3c; }
    QPushButton:pressed { background-color: #1b5e20; }
"""

COMBO_STYLE = """
    QComboBox {
        border: 1px solid #b0bec5;
        padding: 5px 10px;
        border-radius: 4px;
        font-size: 13px;
        min-width: 140px;
    }
"""

LABEL_MUTED = "color: #78909c; font-size: 12px;"
