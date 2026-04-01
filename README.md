# Project Documentation — Passport Data Extractor (OCR)

This repository extracts structured data from **passport images** (MRZ + visual zone) and can also extract fields from a **foreign employment card**-style document. It uses:

- **passporteye** to locate/read the MRZ region
- **EasyOCR** + **Tesseract** to OCR non‑MRZ fields (authority / issue dates / labeled fields)
- **openpyxl** to write results into an Excel file (`PASSPORT-FORM.xlsx`)

---

## Quick start

### Prerequisites

- **Python 3.10+** (recommended: create a virtualenv)
- **Tesseract OCR** installed on your machine (required by `pytesseract`)
  - macOS (Homebrew): `brew install tesseract`
  - Ubuntu/Debian: `sudo apt-get update && sudo apt-get install -y tesseract-ocr`

### Install Python dependencies

```bash
pip install -r requirements.txt
```

Note: the code imports `pytesseract`, `numpy`, and `openpyxl`. If you hit import errors, install them too:

```bash
pip install pytesseract numpy openpyxl
```

---

## Running the script

The primary runnable entrypoint is `passport_data_extractor.py`.

## Desktop app (PySide6)

This repo also includes a desktop UI under `desktop_app/`.

### Run the UI

```bash
python -m desktop_app.main
```

### Build Windows installer (auto installs VC++ runtime)

1. Build app first:

```powershell
pyinstaller --noconfirm "Passport-Data-Extractor.spec"
```

2. Download `vc_redist.x64.exe` from Microsoft and place it at:
`third_party/vc_redist.x64.exe`

3. Install Inno Setup 6.

4. Build installer:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_installer.ps1
```

Installer output will be created in `installer_output/`.

### What the UI supports (v1)

- Load **passport image** + **employment card image**
- Run OCR in the background (UI stays responsive)
- Show a preview with a simple MRZ overlay (estimated region)
- Export extracted data to **Excel**, **CSV**, or **JSON**
- Save a basic local **history** of scans (stored in your OS app-data folder)

### Default run (as currently configured)

```bash
python passport_data_extractor.py
```

By default, the `__main__` block uses:

- `country_codes_file = 'data/country_codes.json'`
- `passport_img = 'images/pass_empoy.jpg'`
- `employee_card_img = 'images/pass_empoy.jpg'`
- `xlsx_path = 'PASSPORT-FORM.xlsx'`

It will:

1. Run extraction with three OCR modes: **EasyOCR**, **Tesseract**, and **Both**
2. Print a side‑by‑side comparison table
3. Save the **Both** result to `PASSPORT-FORM.xlsx`

---

## Inputs and outputs

### Inputs

- **Country codes**: `data/country_codes.json`
  - Used to map MRZ nationality codes to full country names.
- **Images**: put test images in `images/`
  - Sample images are already present (e.g. `Indonesia.jpg`, `India.jpg`, etc.)

### Outputs

- **Console output**:
  - Field values printed via `print_data()` or the comparison table in `get_passport_and_card_data_all_engines()`.
- **Excel output**:
  - `save_to_excel()` appends one row into `PASSPORT-FORM.xlsx` (creates it with headers if missing).

---

## Main API (Python usage)

### Passport extraction only

```python
from passport_data_extractor import PassportDataExtractor

extractor = PassportDataExtractor("data/country_codes.json", gpu=True)
data = extractor.get_data("images/Indonesia.jpg", ocr_engine="both")
extractor.print_data(data)
```

### Passport + card extraction (combined form)

```python
from passport_data_extractor import PassportDataExtractor

extractor = PassportDataExtractor("data/country_codes.json", gpu=True)
combined = extractor.get_passport_and_card_data(
    passport_img_name="images/Indonesia.jpg",
    card_img_name="images/employee.png",
    ocr_engine="both",
)
extractor.save_to_excel(combined, "PASSPORT-FORM.xlsx")
```

### OCR engines

Many methods accept `ocr_engine`:

- `easyocr`: EasyOCR only
- `tesseract`: Tesseract only
- `both`: merges lines from both engines (default in code paths)

---

## How extraction works (high level)

### MRZ-driven fields

`get_data()` calls `passporteye.read_mrz()` to find the MRZ, OCRs its ROI, and parses:

- Name (surname + given names)
- Passport type
- Passport number
- Nationality (country code → country name)
- Date of birth
- Date of expiry

### Visual-zone fields (OCR heuristics)

The full image is OCR’d to find:

- **Authority** via keyword matching (e.g., “ISSUING AUTHORITY”, “ISSUED BY”, …)
- **Date of issue** via:
  - labeled-field matching (various OCR-misspell variants),
  - then keyword proximity search,
  - then a fallback “collect all dates and choose likely issue date” heuristic.

The implementation includes additional cleanup to handle frequent OCR month misreads (e.g. `JAM → JAN`, `JUM → JUN`) and split date tokens across lines.

---

## Notebook

`passport-data-extractor-ocr.ipynb` contains a runnable walkthrough and (in some environments) demonstrates installing Tesseract and Python packages.

---

## Project layout

- `passport_data_extractor.py`: main implementation + runnable example
- `passport-data-extractor-ocr.ipynb`: notebook demo
- `data/`
  - `country_codes.json`: nationality/issuer code mapping
- `images/`: sample passport/card images
- `PASSPORT-FORM.xlsx`: Excel output template/target file
- `OCR Solution for Passport Data Extraction.pptx`: slides/presentation

---

## Troubleshooting

### “TesseractNotFoundError” / empty Tesseract output

- Ensure Tesseract is installed and in `PATH`:
  - macOS: `which tesseract`
  - Linux: `which tesseract`

### GPU / EasyOCR issues

`PassportDataExtractor(..., gpu=True)` enables GPU use for EasyOCR. If you don’t have a compatible setup, use:

```python
extractor = PassportDataExtractor("data/country_codes.json", gpu=False)
```
