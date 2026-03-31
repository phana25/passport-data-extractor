import sys

from PySide6.QtWidgets import QApplication

from desktop_app.ui.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Passport Data Extractor")
    app.setOrganizationName("ByteLab")

    w = MainWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

