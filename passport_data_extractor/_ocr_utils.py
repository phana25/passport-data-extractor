import os
import sys
import shutil
import re
import cv2
import pytesseract


class OCRMixin:
    _MONTH_FIXES = {
        'JAM': 'JAN', 'JAN': 'JAN',
        'FEB': 'FEB',
        'MAR': 'MAR',
        'APR': 'APR',
        'MAY': 'MAY',
        'JUM': 'JUN', 'JUN': 'JUN',
        'JUL': 'JUL',
        'AUG': 'AUG',
        'SEP': 'SEP',
        'OCT': 'OCT',
        'NOV': 'NOV',
        'DEC': 'DEC',
    }

    def _configure_tesseract(self):
        configured = getattr(pytesseract.pytesseract, "tesseract_cmd", "") or ""
        if configured and os.path.exists(configured):
            return True

        exe = shutil.which("tesseract")
        if exe:
            pytesseract.pytesseract.tesseract_cmd = exe
            return True

        candidates = []
        is_windows = sys.platform == "win32"
        binary_name = "tesseract.exe" if is_windows else "tesseract"

        if getattr(sys, "frozen", False):
            exe_dir = os.path.dirname(sys.executable)
            candidates.append(os.path.join(exe_dir, "tesseract", binary_name))
            meipass = getattr(sys, "_MEIPASS", "")
            if meipass:
                candidates.append(os.path.join(meipass, "tesseract", binary_name))

        here = os.path.dirname(os.path.abspath(__file__))
        candidates.append(os.path.join(here, "vendor", "tesseract", binary_name))

        if not is_windows:
            candidates.extend([
                "/opt/homebrew/bin/tesseract",
                "/usr/local/bin/tesseract",
                "/usr/bin/tesseract",
            ])
        else:
            candidates.extend([
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            ])

        for candidate in candidates:
            if os.path.exists(candidate):
                pytesseract.pytesseract.tesseract_cmd = candidate
                tessdata_dir = os.path.join(os.path.dirname(candidate), "tessdata")
                if os.path.isdir(tessdata_dir):
                    os.environ.setdefault("TESSDATA_PREFIX", tessdata_dir)
                return True

        return False

    def _fix_month_typos(self, text):
        text = re.sub(
            r'\d+[^\s/]*/\s*(?=(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC))',
            '', text, flags=re.IGNORECASE
        )
        text = re.sub(
            r'\b\w*?(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\b',
            lambda m: m.group(1).upper(),
            text, flags=re.IGNORECASE
        )

        def _replace(m):
            return self._MONTH_FIXES.get(m.group(0).upper(), m.group(0))
        return re.sub(r'\b(JAM|JUM|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b',
                      _replace, text, flags=re.IGNORECASE)

    def _tesseract_lines(self, img):
        if not self._tesseract_available:
            return []
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (3, 3), 0)
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        try:
            text = pytesseract.image_to_string(thresh, config='--psm 6 --oem 3')
        except pytesseract.pytesseract.TesseractNotFoundError:
            self._tesseract_available = False
            return []
        return [self._fix_month_typos(l.strip()) for l in text.splitlines() if l.strip()]

    def _easyocr_lines(self, img):
        return [self._fix_month_typos(l) for l in self.reader.readtext(img, detail=0)]

    # ------------------------------------------------------------------
    # Combined — deduplicate exact lines from both engines
    # ------------------------------------------------------------------

    def _dual_ocr_lines(self, img):
        seen: set = set()
        out: list = []
        for line in self._easyocr_lines(img) + self._tesseract_lines(img):
            key = re.sub(r'\s+', '', (line or '').upper())
            if key and key not in seen:
                seen.add(key)
                out.append(line)
        return out

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def _ocr_lines(self, img, engine='both'):
        if engine == 'easyocr':
            return self._easyocr_lines(img)
        elif engine == 'tesseract':
            if not self._tesseract_available:
                return self._easyocr_lines(img)
            return self._tesseract_lines(img)
        if engine == 'both' and not self._tesseract_available:
            return self._easyocr_lines(img)
        return self._dual_ocr_lines(img)

    # ------------------------------------------------------------------
    # Text normalisation
    # ------------------------------------------------------------------

    def _normalize_ocr_line(self, line):
        cleaned = re.sub(r'[\t\r\n]+', ' ', line)
        cleaned = re.sub(r'\s{2,}', ' ', cleaned)
        return cleaned.strip()
