import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QFont

from db.database import init_db
from ui.main_window import MainWindow


def main():
    init_db()

    app = QApplication(sys.argv)
    app.setApplicationName("Kaizen Drilling Invoice Manager")
    app.setFont(QFont("Segoe UI", 10))

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
