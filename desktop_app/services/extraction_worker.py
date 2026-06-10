from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import logging
import sys
import pytesseract

from PySide6.QtCore import QObject, Signal

from passport_data_extractor import PassportDataExtractor

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class ScanResult:
    passport_data: dict
    card_data: dict
    combined: dict


class ExtractionWorker(QObject):
    progress = Signal(int)  # 0..100
    status = Signal(str)
    finished = Signal(object)  # ScanResult
    failed = Signal(str)
    extractor_ready = Signal(object)

    def __init__(
        self,
        country_codes_file: str,
        passport_path: str, 
        card_path: str,
        ocr_engine: str = "both",
        gpu: bool = True,
        extractor: PassportDataExtractor | None = None,
    ) -> None:
        super().__init__()
        self.country_codes_file = country_codes_file
        self.passport_path = passport_path
        self.card_path = card_path
        self.ocr_engine = ocr_engine
        self.gpu = gpu
        self._extractor_provided = extractor

    def run(self) -> None:
        try:
            self.progress.emit(2)
            logger.info(f"Scan started — passport={self.passport_path!r} card={self.card_path!r} engine={self.ocr_engine} gpu={self.gpu}")

            if self._extractor_provided:
                self.status.emit("Using cached OCR engine…")
                extractor = self._extractor_provided
            else:
                self.status.emit("Initializing OCR engine (first time)…")
                extractor = PassportDataExtractor(self.country_codes_file, gpu=self.gpu)
                self.extractor_ready.emit(extractor)

            if not extractor._tesseract_available:
                self.status.emit("OCR: Tesseract missing (using EasyOCR only)")
                logger.warning("Tesseract not available — using EasyOCR only")
            else:
                self.status.emit("OCR: Tesseract engine ready")
                logger.info("Tesseract ready")

            passport_data: dict = {}
            if self.passport_path:
                self.status.emit("Reading Passport (Step 1: Reading MRZ)…")
                self.progress.emit(10)
                passport_data = extractor.get_data(self.passport_path, ocr_engine=self.ocr_engine) or {}
                self.progress.emit(45)
                logger.info(f"Passport result: Surname={passport_data.get('Surname')!r} "
                            f"Given={passport_data.get('Given Names')!r} "
                            f"Nationality={passport_data.get('Nationality')!r} "
                            f"Passport#={passport_data.get('Passport Number')!r} "
                            f"DOB={passport_data.get('Date of Birth')!r} "
                            f"Expiry={passport_data.get('Date of Expiry')!r}")

            card_data: dict = {}
            if self.card_path:
                self.status.emit("Reading Employee Card (Step 2: Detailed OCR)…")
                self.progress.emit(55)
                card_data = extractor.get_foreign_employment_card_data(
                    self.card_path, ocr_engine=self.ocr_engine
                ) or {}
                self.progress.emit(85)
                logger.info(f"Card result: {card_data}")

            self.status.emit("Finalizing data and applying smart logic…")
            self.progress.emit(90)

            card_fields = [
                'Card Number', 'DC1', 'DC2', 'DC3',
                'Company Card', 'Position Card'
            ]
            for field in card_fields:
                if field in passport_data and passport_data[field] != 'Not Found':
                    if card_data.get(field) in (None, 'Not Found'):
                        card_data[field] = passport_data[field]

            combined = extractor._build_combined(passport_data, card_data)

            self.progress.emit(100)
            self.status.emit("Done.")
            logger.info("Scan completed successfully")
            try:
                from desktop_app.services.telegram_reporter import report_scan
                report_scan(passport_data, card_data, self.passport_path or self.card_path or "")
            except ImportError:
                pass
            self.finished.emit(ScanResult(passport_data=passport_data, card_data=card_data, combined=combined))
        except Exception as e:  # noqa: BLE001
            logger.exception(f"Scan failed: {e}")
            try:
                from desktop_app.services.telegram_reporter import report_error
                report_error(f"{type(e).__name__}: {e}", self.passport_path or self.card_path or "")
            except ImportError:
                pass
            self.failed.emit(f"{type(e).__name__}: {e}")


def default_country_codes_path() -> str:
    # 1) PyInstaller one-file temp extraction dir
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        p = Path(meipass) / "data" / "country_codes.json"
        if p.exists():
            return str(p)

    # 2) Next to executable (PyInstaller one-folder)
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        p = exe_dir / "data" / "country_codes.json"
        if p.exists():
            return str(p)

    # 3) Project root when running from source
    project_root = Path(__file__).resolve().parents[2]
    p = project_root / "data" / "country_codes.json"
    if p.exists():
        return str(p)

    # 4) Last-resort fallback (original behavior)
    return str(Path("data") / "country_codes.json")

