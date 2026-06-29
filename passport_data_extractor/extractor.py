import os
import ssl
import sys
import re
import logging
import json
import tempfile
import warnings
import threading

logger = logging.getLogger(__name__)

import cv2
import easyocr
import matplotlib.image as mpimg
import pytesseract
import certifi
from passporteye import read_mrz

try:
    from mrzscanner import MRZScanner as _MRZScanner, ModelType as _MRZModelType
    _MRZSCANNER_AVAILABLE = True
except ImportError:
    _MRZSCANNER_AVAILABLE = False
    _MRZModelType = None

try:
    from docaligner import DocAligner as _DocAligner, ModelType as _DocAlignerModelType
    _DOCALIGNER_AVAILABLE = True
except ImportError:
    _DOCALIGNER_AVAILABLE = False
    _DocAligner = None
    _DocAlignerModelType = None

try:
    from fastmrz import FastMRZ as _FastMRZ
    _FASTMRZ_AVAILABLE = True
except ImportError:
    _FASTMRZ_AVAILABLE = False
    _FastMRZ = None

try:
    import requests as _requests
    _REQUESTS_AVAILABLE = True
except ImportError:
    _REQUESTS_AVAILABLE = False

_CLOUD_API_URL = "https://phana25-mrz-scanner-api.hf.space/scan"

warnings.filterwarnings('ignore')

from ._ocr_utils import OCRMixin
from ._mrz_parser import MRZMixin
from ._field_extractor import FieldExtractorMixin
from ._excel_writer import ExcelMixin


def _configure_ssl_certificates():
    cafile = certifi.where()
    os.environ.setdefault("SSL_CERT_FILE", cafile)
    os.environ.setdefault("REQUESTS_CA_BUNDLE", cafile)
    ssl._create_default_https_context = lambda: ssl.create_default_context(cafile=cafile)


class PassportDataExtractor(OCRMixin, MRZMixin, FieldExtractorMixin, ExcelMixin):

    def __init__(self, country_codes_file, gpu=True):
        _configure_ssl_certificates()
        self._tesseract_available = self._configure_tesseract()

        if sys.platform == "win32" and gpu:
            try:
                self.reader = easyocr.Reader(lang_list=['en'], gpu=True)
            except Exception:
                self.reader = easyocr.Reader(lang_list=['en'], gpu=False)
        else:
            self.reader = easyocr.Reader(lang_list=['en'], gpu=False)

        self._mrz_scanner = None
        if _MRZSCANNER_AVAILABLE:
            try:
                self._mrz_scanner = _MRZScanner(
                    model_type=_MRZModelType.two_stage,
                    detection_cfg='20250222',
                    recognition_cfg='20250221',
                )
            except Exception:
                try:
                    self._mrz_scanner = _MRZScanner()
                except Exception:
                    pass

        self._doc_aligner = None
        if _DOCALIGNER_AVAILABLE:
            try:
                self._doc_aligner = _DocAligner(
                    model_type=_DocAlignerModelType.heatmap,
                )
            except Exception:
                pass

        self._fast_mrz = None
        if _FASTMRZ_AVAILABLE:
            try:
                self._fast_mrz = _FastMRZ()
            except Exception:
                pass

        with open(country_codes_file) as f:
            self.country_codes = json.load(f)

        if _REQUESTS_AVAILABLE:
            self._start_cloud_keepalive()

    def _start_cloud_keepalive(self):
        """Ping the cloud API immediately on startup and every 10 minutes to prevent HF Space from sleeping."""
        health_url = _CLOUD_API_URL.replace("/scan", "/health")

        def ping():
            while True:
                try:
                    _requests.get(health_url, timeout=5)
                except Exception:
                    pass
                threading.Event().wait(600)

        t = threading.Thread(target=ping, daemon=True)
        t.start()

    def _align_document(self, img):
        """Use DocAligner to perspective-correct a document image.

        Returns the warped (flattened) image, or None if alignment fails.
        DocAligner returns 4 corner points; we apply a perspective warp to
        produce a flat frontal view of the document.
        """
        if self._doc_aligner is None or img is None:
            return None
        try:
            import numpy as np
            corners = self._doc_aligner(img, do_center_crop=False)
            if corners is None or len(corners) != 4:
                return None
            corners = np.array(corners, dtype=np.float32)
            # Compute output dimensions from the detected corners
            w1 = np.linalg.norm(corners[1] - corners[0])
            w2 = np.linalg.norm(corners[2] - corners[3])
            h1 = np.linalg.norm(corners[3] - corners[0])
            h2 = np.linalg.norm(corners[2] - corners[1])
            out_w = int(max(w1, w2))
            out_h = int(max(h1, h2))
            if out_w < 50 or out_h < 50:
                return None
            dst = np.array([
                [0, 0],
                [out_w - 1, 0],
                [out_w - 1, out_h - 1],
                [0, out_h - 1],
            ], dtype=np.float32)
            M = cv2.getPerspectiveTransform(corners, dst)
            warped = cv2.warpPerspective(img, M, (out_w, out_h))
            return warped
        except Exception:
            return None

    def _scan_mrz_with_docsaid(self, img_path):
        """Run MRZScanner (two-stage DL model) on a full image.

        Tries the original image first. If the result is empty or a garbage
        name line (no '<<'), also tries a DocAligner-corrected version and
        keeps whichever produces more valid MRZ lines.

        Returns (mrz_lines, mrz_polygon) where mrz_lines is a list of raw MRZ
        text strings and mrz_polygon is the detected region (or None).
        Returns ([], None) on failure.
        """
        if self._mrz_scanner is None:
            return [], None
        img = cv2.imread(img_path)
        if img is None:
            return [], None
        try:
            result = self._mrz_scanner(
                img=img,
                do_center_crop=False,
                do_postprocess=True,
            )
            texts = result.get('mrz_texts') or []
            polygon = result.get('mrz_polygon')
            lines = [t for t in texts if t and len(t.strip()) >= 20]
            # Check if the name line looks valid (needs '<<' as surname/given separator)
            name_line_ok = any('<<' in t for t in lines)
            if not name_line_ok:
                # Try DocAligner-corrected image
                aligned = self._align_document(img)
                if aligned is not None:
                    try:
                        aligned_result = self._mrz_scanner(
                            img=aligned,
                            do_center_crop=False,
                            do_postprocess=True,
                        )
                        aligned_texts = aligned_result.get('mrz_texts') or []
                        aligned_lines = [t for t in aligned_texts if t and len(t.strip()) >= 20]
                        if any('<<' in t for t in aligned_lines):
                            return aligned_lines, aligned_result.get('mrz_polygon')
                    except Exception:
                        pass
            return lines, polygon
        except Exception:
            return [], None

    def _scan_mrz_cloud(self, img_path):
        """Send image to the cloud MRZScanner API as fallback when local scan is weak."""
        if not _REQUESTS_AVAILABLE:
            return []
        try:
            with open(img_path, 'rb') as f:
                response = _requests.post(
                    _CLOUD_API_URL,
                    files={"file": f},
                    timeout=10,
                )
            if response.status_code == 200:
                return response.json().get("mrz_texts") or []
        except Exception:
            pass
        return []

    def clean(self, string):
        return ''.join(char for char in string if char.isalnum()).upper()

    def get_country_name(self, country_code):
        for country in self.country_codes:
            if country['code'] == country_code:
                return country['name']
        return country_code

    def print_data(self, data):
        for key, value in data.items():
            print(f'{key.replace("_", " ").capitalize()}\t:\t{value}')

    def _split_name(self, full_name):
        if not full_name or full_name == 'Not Found':
            return 'Not Found', 'Not Found'
        parts = [p for p in full_name.strip().split(' ') if p]
        if not parts:
            return 'Not Found', 'Not Found'
        surname = parts[-1]
        given = ' '.join(parts[:-1]) if len(parts) > 1 else 'Not Found'
        return surname, given

    def _normalize_gender(self, raw: str) -> str:
        g = (raw or '').upper().strip()
        if g in ('F', 'FEMALE') or g.startswith('F'):
            return 'F'
        if g in ('M', 'MALE') or g.startswith('M'):
            return 'M'
        return ''

    def _build_combined(self, passport_data, card_data):
        surname = passport_data.get('Surname', 'Not Found')
        given = passport_data.get('Given Names', 'Not Found')
        if surname == 'Not Found' and given == 'Not Found':
            surname, given = self._split_name(passport_data.get('Name', 'Not Found'))
        bd1, bd2, bd3 = self._split_date_components(passport_data.get('Date of Birth', 'Not Found'))
        iss1, iss2, iss3 = self._split_date_components(passport_data.get('Date of Issue', 'Not Found'))
        ed1, ed2, ed3 = self._split_date_components(passport_data.get('Date of Expiry', 'Not Found'))
        full_name_parts = []
        if surname and surname != 'Not Found':
            full_name_parts.append(surname)
        if given and given != 'Not Found':
            full_name_parts.append(given)
        full_name = ' '.join(full_name_parts) if full_name_parts else 'Not Found'
        return {
            'SURNAME':       surname,
            'GSURNAME':      given,
            'BD1': bd1, 'BD2': bd2, 'BD3': bd3,
            'NASTIONALTY':   passport_data.get('Nationality', 'Not Found'),
            'PASSPORT':      passport_data.get('Passport Number', 'Not Found'),
            'ISS1': iss1, 'ISS2': iss2, 'ISS3': iss3,
            'ED1': ed1, 'ED2': ed2, 'ED3': ed3,
            'CARD NUMBER':   card_data.get('Card Number', 'Not Found'),
            'DC1':           card_data.get('DC1', 'Not Found'),
            'DC2':           card_data.get('DC2', 'Not Found'),
            'DC3':           card_data.get('DC3', 'Not Found'),
            'COMPANY CARD':  card_data.get('Company Card', 'Not Found'),
            'POSITOIN CARD': card_data.get('Position Card', 'Not Found'),
            'NAME 02':       full_name,
            'GENDER_RAW':    passport_data.get('Gender', ''),
            'Gender':        self._normalize_gender(passport_data.get('Gender', '')),
        }

    def _extract_fields_from_best_mrz(self, best, mrz_type, ocr_results, ocr_extended, debug=False):
        info = {}
        name = best['given']
        surname = best['surname']

        if mrz_type == 'TD1':
            l1, l2 = best['line1'], best['line2']
            dob = self.parse_birth_date(l2[0:6])
            expiry = self.parse_date(l2[8:14])
            info['Nationality'] = self.get_country_name(self.clean(l2[15:18]))
            info['Passport Type'] = self.clean(l1[0:2])
            info['Passport Number'] = self.clean(l1[5:14])
            sex_ch = l2[7] if len(l2) > 7 else ''
        else:
            a, b = best['line1'], best['line2']
            dob = self.parse_birth_date(b[13:19])
            expiry = self.parse_date(b[21:27])
            info['Nationality'] = self.get_country_name(self.clean(b[10:13]))
            info['Passport Type'] = self.clean(a[0:2])
            info['Passport Number'] = self.clean(b[0:9])
            sex_ch = b[20] if len(b) > 20 else ''

        info['Date of Birth'] = dob
        info['Date of Expiry'] = expiry
        info['Gender'] = 'Male' if sex_ch == 'M' else ('Female' if sex_ch == 'F' else 'Not Found')

        def _best_of(candidates_list):
            best_g, best_s, best_q = '', '', -999
            for g, s in candidates_list:
                if g or s:
                    q = self._mrz_name_quality(g, s)
                    if g and s:  # prefer complete pairs over partial ones
                        q += 10
                    if not self._is_suspicious_name(g, s) and q > best_q:
                        best_g, best_s, best_q = g, s, q
            return best_g, best_s

        v_given, v_surname = self._extract_visual_latin_name(ocr_results)
        lbl_given, lbl_surname = self._extract_name_from_visual_labels(ocr_results)
        ref_sn = surname or v_surname or lbl_surname
        given_from_sn = self._extract_given_by_surname_from_visual(ocr_results, ref_sn) if ref_sn else ''

        # Strip known 2-char OCR filler tokens from visual given names so they
        # don't beat a valid MRZ result purely on letter-count arithmetic.
        _FILLER_GIVENS = {
            'EY', 'YE', 'YK', 'KY', 'EK', 'KE', 'SK', 'KS', 'KC', 'CK',
            'ES', 'KX', 'XK', 'EX', 'XE', 'FO', 'OF', 'FE', 'EF', 'CE', 'EC',
        }
        def _strip_filler(g):
            toks = [t for t in (g or '').split() if t.upper() not in _FILLER_GIVENS]
            return ' '.join(toks)
        v_given = _strip_filler(v_given)
        lbl_given = _strip_filler(lbl_given)
        given_from_sn = _strip_filler(given_from_sn)

        # If MRZ gave a clean complete result, trust it — don't let visual OCR override it.
        mrz_clean = (name and surname
                     and not self._is_suspicious_name(name, surname))
        if not mrz_clean:
            best_vg, best_vs = _best_of([
                (name, surname),
                (v_given, v_surname),
                (lbl_given, lbl_surname),
                (given_from_sn, ref_sn) if given_from_sn else ('', ''),
            ])
            if best_vg or best_vs:
                if self._mrz_name_quality(best_vg, best_vs) > self._mrz_name_quality(name, surname) \
                        or self._is_suspicious_name(name, surname) \
                        or (not name and best_vg
                            and max((len(t) for t in best_vg.split()), default=0) >= 3):
                    name, surname = best_vg, best_vs
        if self._is_suspicious_name(name, surname) or not name:
            fg, fs = self._extract_mrz_name_from_ocr_lines(ocr_results)
            if fg and fs:
                name, surname = fg, fs
            elif v_given and v_surname and not self._is_suspicious_name(v_given, v_surname):
                name, surname = v_given, v_surname

        # Cross-validate given name against visual reference to strip OCR artifacts
        for ref in (given_from_sn, lbl_given):
            if not ref or not name or name == ref:
                continue
            name_toks = name.split()
            ref_toks = ref.split()
            # Case 1: trailing short filler tokens (e.g. "THI HUONG ES" → "THI HUONG")
            if (len(name_toks) > len(ref_toks)
                    and name_toks[:len(ref_toks)] == ref_toks
                    and all(len(t) <= 2 for t in name_toks[len(ref_toks):])):
                name = ref
                break
            # Case 2: leading OCR artifact char in a token (e.g. "VAN KDUY"→"VAN DUY", "NHU XPHUONG"→"NHU PHUONG")
            if len(name_toks) == len(ref_toks) and len(name_toks) >= 2:
                corrected = []
                for nt, rt in zip(name_toks, ref_toks):
                    if (nt != rt and len(nt) > len(rt)
                            and nt[0] in ('K', 'X') and nt[1:] == rt and len(rt) >= 2):
                        corrected.append(rt)
                    else:
                        corrected.append(nt)
                if corrected == ref_toks:
                    name = ref
                    break
            # Case 3: trailing single-char artifact on last token (e.g. "BINHK"→"BINH", "ANHK"→"ANH")
            if len(name_toks) == len(ref_toks):
                corrected = []
                for nt, rt in zip(name_toks, ref_toks):
                    if (nt != rt and nt.startswith(rt)
                            and len(nt) - len(rt) == 1
                            and nt[-1] in ('K', 'X')
                            and len(rt) >= 2):
                        corrected.append(rt)
                    else:
                        corrected.append(nt)
                if corrected == ref_toks:
                    name = ref
                    break

        info['Given Names'] = name or 'Not Found'
        info['Surname'] = surname or 'Not Found'
        info['Name'] = ' '.join(p for p in [info['Surname'], info['Given Names']] if p != 'Not Found') or 'Not Found'

        if debug:
            print(f'DEBUG MRZ type: {mrz_type}')
            print(f'DEBUG MRZ score: {best.get("score")}')
            print(f'DEBUG MRZ checks: {best.get("details")}')

        return info, dob, expiry

    def _get_data_from_ocr_fallback(self, img_name, ocr_engine='both', debug=False):
        full_img = cv2.imread(img_name)
        if full_img is None:
            return {}
        ocr_results = self._ocr_lines(full_img, engine=ocr_engine)
        ocr_extended = self._rejoin_split_ocr_dates(ocr_results)

        best_td3 = self._select_best_mrz_candidate(ocr_results)
        best_td1 = self._select_best_td1_candidate(ocr_results)

        td3_score = best_td3['score'] if best_td3 else -9999
        td1_score = best_td1['score'] if best_td1 else -9999

        if td3_score <= -50 and td1_score <= -50:
            return {}

        if td3_score >= td1_score and best_td3:
            best, mrz_type = best_td3, 'TD3'
        else:
            best, mrz_type = best_td1, 'TD1'

        info, dob, expiry = self._extract_fields_from_best_mrz(
            best, mrz_type, ocr_results, ocr_extended, debug=debug)

        issue_date = self._extract_date_for_labels(ocr_extended, [
            'DATE OF ISSUE', 'DATE OF ISSUANCE', 'ISSUED ON', 'ISSUED DATE', 'ISSUE DATE'])
        if issue_date == 'Not Found':
            issue_date = self.find_issuing_date(ocr_extended, dob_str=dob, expiry_str=expiry)
        info['Date of Issue'] = issue_date
        info['Authority'] = self.find_authority(ocr_results)
        return info

    def get_data(self, img_name, debug=False, ocr_engine='both'):
        user_info = {}
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmpfile:
            tmpfile_path = tmpfile.name

        try:
            # --- Primary: Cloud API ---
            cloud_lines = self._scan_mrz_cloud(img_name)
            if debug and cloud_lines:
                print(f'DEBUG Cloud MRZ lines: {cloud_lines}')
            cloud_good = len(cloud_lines) >= 2 and any('<<' in l for l in cloud_lines)

            # If cloud succeeded, skip local MRZScanner entirely.
            if cloud_good:
                dl_lines = cloud_lines
                dl_polygon = None
            else:
                # --- Fallback: local MRZScanner (deep learning) ---
                dl_lines, dl_polygon = self._scan_mrz_with_docsaid(img_name)
                if debug and dl_lines:
                    print(f'DEBUG MRZScanner (DL) lines: {dl_lines}')

            dl_good = len(dl_lines) >= 2 and any('<<' in l for l in dl_lines)

            mrz = None
            if not dl_good:
                try:
                    mrz = read_mrz(img_name, save_roi=True)
                except (pytesseract.pytesseract.TesseractNotFoundError, FileNotFoundError, OSError):
                    self._tesseract_available = False

            if dl_lines or mrz:
                # Build the initial code list by pooling lines from both sources.
                code = list(dl_lines)
                if mrz:
                    mrz_roi = mrz.aux['roi']
                    mpimg.imsave(tmpfile_path, mrz_roi, cmap='gray')
                    roi_img = cv2.imread(tmpfile_path)
                    roi_code = self._ocr_mrz_roi(roi_img)
                    code += roi_code
                    if debug:
                        print(f'DEBUG passporteye ROI lines: {roi_code}')

                if len(code) < 2:
                    # Neither MRZScanner nor passporteye produced enough MRZ lines.
                    # Fall back to full-image OCR which searches the whole page for MRZ patterns.
                    print(f'Warning: Insufficient MRZ lines for {img_name}, trying full-image OCR fallback.')
                    fallback = self._get_data_from_ocr_fallback(img_name, ocr_engine=ocr_engine, debug=debug)
                    if fallback:
                        user_info.update(fallback)
                    else:
                        print(f'OCR fallback also failed for {img_name}.')
                    return user_info

                td1_count = sum(1 for l in code if 26 <= len(re.sub(r'\s', '', l)) <= 34)
                td3_count = sum(1 for l in code if 40 <= len(re.sub(r'\s', '', l)) <= 48)
                mrz_type_hint = getattr(mrz, 'type', None) or ('TD1' if td1_count > td3_count else 'TD3')

                if mrz_type_hint == 'TD1':
                    best_mrz = self._select_best_td1_candidate(code)
                    if not best_mrz:
                        best_mrz = self._select_best_mrz_candidate(code, mrz_obj=mrz)
                        mrz_type_hint = 'TD3'
                else:
                    best_mrz = self._select_best_mrz_candidate(code, mrz_obj=mrz)
                    if not best_mrz or best_mrz['score'] < -30:
                        td1_candidate = self._select_best_td1_candidate(code)
                        if td1_candidate and (not best_mrz or td1_candidate['score'] > best_mrz['score']):
                            best_mrz = td1_candidate
                            mrz_type_hint = 'TD1'

                if not best_mrz:
                    return print(f'Error: Unable to build MRZ candidates for image {img_name}.')

                # Score DL lines in isolation so ROI garbage cannot contaminate
                # the trust baseline. Only promote DL-only best if it has a valid name
                # (prevents garbage DL lines that score well on check digits alone from
                # suppressing a correct passporteye result).
                dl_only_best = None
                if dl_lines and len(dl_lines) >= 2:
                    if mrz_type_hint == 'TD1':
                        dl_only_best = self._select_best_td1_candidate(dl_lines)
                    else:
                        dl_only_best = self._select_best_mrz_candidate(dl_lines)
                    if dl_only_best and dl_only_best['score'] >= 30:
                        dl_sur = dl_only_best.get('surname', '') or ''
                        dl_giv = dl_only_best.get('given', '') or ''
                        if (dl_sur or dl_giv) and not self._is_suspicious_name(dl_giv, dl_sur):
                            best_mrz = dl_only_best

                full_img = cv2.imread(img_name)
                if debug and full_img is not None:
                    if dl_polygon is not None:
                        # Draw DL polygon from MRZScanner
                        import numpy as np
                        pts = np.array(dl_polygon, dtype=int).reshape((-1, 1, 2))
                        debug_img = full_img.copy()
                        cv2.polylines(debug_img, [pts], True, (0, 255, 0), 2)
                        cv2.putText(
                            debug_img, 'MRZ (DL)', tuple(pts[0][0]),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2, cv2.LINE_AA
                        )
                        base, ext = os.path.splitext(img_name)
                        debug_path = f'{base}_mrz_box{ext if ext else ".png"}'
                        cv2.imwrite(debug_path, debug_img)
                        print(f'DEBUG MRZ box saved: {debug_path}')
                    elif mrz:
                        mrz_roi_for_debug = mrz.aux['roi']
                        bbox = self._locate_roi_in_full_image(full_img, mrz_roi_for_debug)
                        if bbox:
                            x, y, w, h = bbox
                            debug_img = full_img.copy()
                            cv2.rectangle(debug_img, (x, y), (x + w, y + h), (0, 255, 0), 2)
                            cv2.putText(
                                debug_img, 'MRZ ROI', (x, max(20, y - 8)),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2, cv2.LINE_AA
                            )
                            base, ext = os.path.splitext(img_name)
                            debug_path = f'{base}_mrz_box{ext if ext else ".png"}'
                            cv2.imwrite(debug_path, debug_img)
                            print(f'DEBUG MRZ box saved: {debug_path}')

                ocr_results = self._ocr_lines(full_img, engine=ocr_engine)
                ocr_extended = self._rejoin_split_ocr_dates(ocr_results)

                # Re-select best MRZ using combined ROI + full-image OCR lines.
                # Full-image OCR sometimes reads `<` chars more faithfully than the
                # aggressively preprocessed ROI, so the clean line can win here.
                combined_code = code + ocr_results
                if mrz_type_hint == 'TD1':
                    full_candidate = self._select_best_td1_candidate(combined_code)
                else:
                    full_candidate = self._select_best_mrz_candidate(combined_code, mrz_obj=(mrz if not dl_lines else None))
                # When DL scanner gave good lines (high score), require a substantial improvement
                # before allowing the combined-code candidate to override, so that OCR noise
                # (e.g., issuing-city text concatenated with the MRZ) cannot displace a clean result.
                # Use dl_only_best score (not combined best_mrz) so ROI garbage doesn't
                # inflate the threshold and trick dl_trusted into seeing a "trusted" garbage result.
                # dl_only_score is only meaningful when dl_only_best was actually promoted
                # as best_mrz (i.e., it has a valid name). If it was rejected (empty/suspicious
                # name), don't apply a trust margin — let combined_code compete freely.
                dl_only_promoted = (dl_only_best is not None and best_mrz is dl_only_best)
                dl_only_score = dl_only_best['score'] if dl_only_promoted else -9999
                dl_trusted = dl_lines and dl_only_score >= 30
                override_margin = 20 if dl_trusted else 0
                if full_candidate and full_candidate['score'] >= best_mrz['score'] + override_margin:
                    best_mrz = full_candidate

                extracted, dob, expiry = self._extract_fields_from_best_mrz(
                    best_mrz, mrz_type_hint, ocr_results, ocr_extended, debug=debug)
                user_info.update(extracted)
                dob = user_info.get('Date of Birth')
                expiry = user_info.get('Date of Expiry')

                issue_date = self._extract_date_for_labels(
                    ocr_extended,
                    ['DATE OF ISSUE', 'DATE OF ISSUANCE', 'ISSUED ON', 'ISSUED DATE', 'ISSUE DATE',
                     'Date of Issue', 'Dale pf issue', 'Dale of issue']
                )
                if issue_date == 'Not Found':
                    issue_date = self._extract_date_near_keywords(
                        ocr_extended,
                        ['DATE OF ISSUE', 'DATE OF ISTUE', 'DALE PF ISSUE', 'DALE OF ISSUE', 'ISSUANCE']
                    )
                if issue_date == 'Not Found':
                    issue_date = self.find_issuing_date(
                        self._rejoin_split_ocr_dates(ocr_results), dob_str=dob, expiry_str=expiry)
                if debug:
                    print('DEBUG full OCR lines:')
                    for line in ocr_results:
                        print(f'  {line!r}')
                    print('DEBUG all dates found:', self._collect_all_dates(ocr_results))
                user_info['Date of Issue'] = issue_date

                if user_info.get('Gender') == 'Not Found':
                    sex = (getattr(mrz, 'sex', '') or '').upper()
                    if sex in ('M', 'F'):
                        user_info['Gender'] = 'Male' if sex == 'M' else 'Female'
                if user_info.get('Gender') == 'Not Found':
                    gender_results = self._extract_labeled_fields(ocr_results, {'Gender': ['GENDER']})
                    raw_g = gender_results.get('Gender', 'Not Found')
                    if raw_g != 'Not Found':
                        user_info['Gender'] = 'Male' if 'M' in raw_g.upper() else (
                            'Female' if 'F' in raw_g.upper() else 'Not Found')

                user_info['Authority'] = self.find_authority(ocr_results)

                combined_text = ' '.join(ocr_results).upper()
                card_keywords = ['FOREIGN EMPLOYMENT CARD', 'EMPLOYMENT CARD', 'FWCMS']
                is_card_detected = any(kw in combined_text for kw in card_keywords)
                if debug:
                    print(f'DEBUG: Combined card detected: {is_card_detected}')
                if is_card_detected:
                    card_data = self.get_foreign_employment_card_data(img_name, ocr_engine=ocr_engine)
                    if debug:
                        print(f'DEBUG: Extracted card data: {card_data}')
                    user_info.update(card_data)

            else:
                logger.warning(f'Machine cannot read MRZ from image {img_name}. Trying OCR fallback…')
                fallback = self._get_data_from_ocr_fallback(img_name, ocr_engine=ocr_engine, debug=debug)
                if fallback:
                    user_info.update(fallback)
                else:
                    logger.warning(f'OCR fallback also failed for {img_name}.')

        finally:
            if os.path.exists(tmpfile_path):
                os.remove(tmpfile_path)

        return user_info

    def get_foreign_employment_card_data(self, img_name, ocr_engine='both'):
        full_img = cv2.imread(img_name)
        if full_img is None:
            print(f'Image not found or unreadable: {img_name}')
            return {}

        ocr_results = self._ocr_lines(full_img, engine=ocr_engine)

        label_map = {
            'Company': ['COMPANY', 'COMPANY NAME', 'ENTERPRISE NAME'],
            'Position': ['POSITION', 'POSTION', 'POSITOIN'],
            'Card Number': ['CARD NUMBER', 'CARD NO', 'ID NO', 'ID NO.', 'I0 NO', 'I0 NO.'],
            'DC1': ['DC1', 'EXPIRED DATE', 'EXPIRY DATE', 'DATE OF EXPIRY'],
            'DC2': ['DC2'],
            'DC3': ['DC3'],
            'Company Card': ['COMPANY CARD'],
            'Position Card': ['POSITION CARD', 'POSITOIN CARD'],
            'Phone': ['PHONE', 'TEL', 'MOBILE'],
            'D01': ['D01'],
            'D02': ['D02'],
            'D03': ['D03'],
            'Name 02': ['NAME 02', 'NAME02'],
        }

        results = self._extract_labeled_fields(ocr_results, label_map)

        expired_date = self._extract_date_for_label(ocr_results, 'Expired Date')
        if expired_date == 'Not Found':
            expired_date = results.get('DC1', 'Not Found')
        day, month, year = self._split_date_components(expired_date)

        if day != 'Not Found':
            results['DC1'] = day
        if month != 'Not Found':
            results['DC2'] = month
        if year != 'Not Found':
            results['DC3'] = year

        if results.get('Company Card') == 'Not Found':
            results['Company Card'] = results.get('Company', 'Not Found')
        if results.get('Position Card') == 'Not Found':
            results['Position Card'] = results.get('Position', 'Not Found')

        return results

    def get_passport_and_card_data(self, passport_img_name, card_img_name,
                                   debug=False, ocr_engine='both'):
        passport_data = self.get_data(passport_img_name, debug=debug, ocr_engine=ocr_engine)
        card_data = self.get_foreign_employment_card_data(card_img_name, ocr_engine=ocr_engine)
        return self._build_combined(passport_data, card_data)

    def get_passport_and_card_data_all_engines(self, passport_img_name, card_img_name, debug=False):
        results = {}
        for engine in ('easyocr', 'tesseract', 'both'):
            passport_data = self.get_data(passport_img_name, debug=debug, ocr_engine=engine)
            card_data = self.get_foreign_employment_card_data(card_img_name, ocr_engine=engine)
            results[engine] = self._build_combined(passport_data, card_data)

        fields = list(results['easyocr'].keys())
        col_w = 22
        header = f"{'FIELD':<18}  {'EASYOCR':<{col_w}}  {'TESSERACT':<{col_w}}  {'BOTH':<{col_w}}"
        print(header)
        print('-' * len(header))
        for field in fields:
            easy = results['easyocr'].get(field, '')
            tess = results['tesseract'].get(field, '')
            both = results['both'].get(field, '')
            print(f"{field:<18}  {easy:<{col_w}}  {tess:<{col_w}}  {both:<{col_w}}")

        return results


if __name__ == "__main__":
    country_codes_file = 'data/country_codes.json'
    passport_img = 'images/pass_empoy.jpg'
    employee_card_img = 'images/pass_empoy.jpg'
    xlsx_path = 'PASSPORT-FORM.xlsx'

    extractor = PassportDataExtractor(country_codes_file)
    results = extractor.get_passport_and_card_data_all_engines(passport_img, employee_card_img)
    extractor.save_to_excel(results['both'], xlsx_path)
