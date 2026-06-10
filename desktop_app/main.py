import sys
import logging
from pathlib import Path

from PySide6.QtWidgets import QApplication

from desktop_app.ui.main_window import MainWindow
from desktop_app import __version__


class _TelegramHandler(logging.Handler):
    """Forwards WARNING+ log records to Telegram."""

    def emit(self, record: logging.LogRecord) -> None:
        try:
            from desktop_app.services.telegram_reporter import send_async
            import platform
            from datetime import datetime

            level_icon = {
                logging.WARNING:  "⚠️",
                logging.ERROR:    "🔴",
                logging.CRITICAL: "🆘",
            }.get(record.levelno, "⚠️")

            ts  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            pc  = platform.node()
            msg = self.format(record)
            text = (
                f"{level_icon} <b>{record.levelname}</b>  v{__version__}\n"
                f"🕐 {ts}  💻 {pc}\n"
                f"<code>{msg[:800]}</code>"
            )
            send_async(text)
        except Exception:
            pass


def _setup_logging() -> Path:
    if getattr(sys, "frozen", False):
        log_dir = Path(sys.executable).resolve().parent / "logs"
    else:
        log_dir = Path(__file__).resolve().parents[1] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "app.log"

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(fmt)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(fmt)

    telegram_handler = _TelegramHandler()
    telegram_handler.setLevel(logging.WARNING)
    telegram_handler.setFormatter(logging.Formatter("%(name)s: %(message)s"))

    root = logging.getLogger()
    root.setLevel(logging.WARNING)
    root.addHandler(file_handler)
    root.addHandler(stream_handler)
    root.addHandler(telegram_handler)

    return log_file


def main() -> int:
    log_file = _setup_logging()
    logger = logging.getLogger("main")

    def handle_exception(exc_type, exc_value, exc_tb):
        logger.critical("Unhandled exception", exc_info=(exc_type, exc_value, exc_tb))
    sys.excepthook = handle_exception

    app = QApplication(sys.argv)
    app.setApplicationName("Passport Data Extractor")
    app.setOrganizationName("ByteLab")

    try:
        from desktop_app.services.telegram_reporter import report_startup
        report_startup()
    except ImportError:
        pass

    w = MainWindow()
    w.show()
    code = app.exec()
    return code


if __name__ == "__main__":
    raise SystemExit(main())
