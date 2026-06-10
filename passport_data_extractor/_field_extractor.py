import re
import datetime as _dt
import cv2
from dateutil import parser


class FieldExtractorMixin:

    def parse_birth_date(self, date_string):
        try:
            date = parser.parse(date_string, yearfirst=True).date()
            if date.year > _dt.date.today().year:
                date = date.replace(year=date.year - 100)
            return date.strftime('%d/%m/%Y')
        except ValueError:
            return None

    def parse_date(self, date_string, dayfirst=False):
        try:
            date = parser.parse(date_string, yearfirst=not dayfirst, dayfirst=dayfirst).date()
            return date.strftime('%d/%m/%Y')
        except (ValueError, TypeError):
            return None

    def _parse_ocr_date(self, date_string):
        cleaned = date_string.strip().upper()
        if re.match(r'^[L|I|S|O|B|0-9].*', cleaned):
            cleaned = re.sub(r'^([L|I])(\d)\b', r'1\2', cleaned)
            cleaned = re.sub(r'^(\d)([L|I])\b', r'\g<1>1', cleaned)
            cleaned = re.sub(r'^(S)(\d)\b', r'5\2', cleaned)
            cleaned = re.sub(r'^(\d)(S)\b', r'\g<1>5', cleaned)
            cleaned = re.sub(r'^(O)(\d)\b', r'0\2', cleaned)
            cleaned = re.sub(r'^(\d)(O)\b', r'\g<1>0', cleaned)
            cleaned = re.sub(r'\b(2[O|0]2[L|I|1])\b', r'2021', cleaned)
            cleaned = re.sub(r'\b(2[O|0]2[O|0])\b', r'2020', cleaned)
            cleaned = re.sub(r'\b2[O|0][L|I|1](\d)\b', r'201\1', cleaned)
            if "2O" in cleaned:
                cleaned = cleaned.replace("2O", "20")
        months = 'JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC'
        m = re.search(r'([0-9]{1,2})[.,\sA-Z/]+(' + months + r')[A-Z]*[.,\s/]+([0-9]{2,4})', cleaned)
        if m:
            cleaned = f"{m.group(1)} {m.group(2)} {m.group(3)}"
        else:
            cleaned = re.sub(r'(\d{1,2})[.,]\s+', r'\1 ', cleaned)
        try:
            date = parser.parse(
                cleaned, dayfirst=True,
                default=_dt.datetime(2000, 1, 1)
            ).date()
            return date.strftime('%d/%m/%Y')
        except (ValueError, TypeError):
            return None

    def _rejoin_split_ocr_dates(self, ocr_lines):
        month_names = 'JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC'
        mon_yr_pat = re.compile(
            r'^(?:' + month_names + r')[A-Z]*\s+\d{4}$', re.IGNORECASE
        )
        day_pat = re.compile(r'^\d{1,2}$')
        dd_mon_pat = re.compile(r'^\d{1,2}\s+(?:' + month_names + r')[A-Z]*$', re.IGNORECASE)
        normalized = [self._normalize_ocr_line(l) for l in ocr_lines if l.strip()]
        rejoined = list(normalized)
        for i, line in enumerate(normalized):
            if mon_yr_pat.match(line):
                for back in range(1, 6):
                    if i - back < 0:
                        break
                    prev = normalized[i - back].strip()
                    if day_pat.match(prev):
                        rejoined.append(f'{prev} {line}')
                        break
            if dd_mon_pat.match(line):
                if i + 1 < len(normalized) and re.fullmatch(r'\d{4}', normalized[i + 1].strip()):
                    rejoined.append(f'{line} {normalized[i + 1].strip()}')
            if i + 1 < len(normalized):
                nxt = normalized[i + 1].strip()
                if day_pat.match(line.strip()) and mon_yr_pat.match(nxt):
                    rejoined.append(f'{line.strip()} {nxt}')
        return rejoined

    def _get_all_date_patterns(self):
        d = r'[0-9OISlB]'
        months = r'JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC'
        return [
            f'{d}{{2}}[-/.]{d}{{2}}[-/.]{d}{{2,4}}',
            f'{d}{{4}}[-/.]{d}{{2}}[-/.]{d}{{2}}',
            f'{d}{{1,2}} {d}{{1,2}} {d}{{2,4}}',
            f'{d}{{1,2}}[.,\\sA-Z/]+(?:{months})[A-Z]*[.,\\s/]+{d}{{2,4}}',
            f'\\b(?:{months})\\b[.,\\sA-Z/]+{d}{{1,2}}[.,\\sA-Z/]+{d}{{2,4}}',
            f'\\b(?:{months})\\b\\s+{d}{{2,4}}',
        ]

    def _collect_all_dates(self, ocr_text):
        mon_only_pat = re.compile(
            r'^(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*\s+\d{4}$', re.IGNORECASE
        )
        full_text_blob = " ".join(ocr_text)
        search_lines = list(ocr_text) + [full_text_blob]
        best = {}
        for line in search_lines:
            for pattern in self._get_all_date_patterns():
                for date_match in re.findall(pattern, line, re.IGNORECASE):
                    parsed = self._parse_ocr_date(date_match)
                    if not parsed:
                        continue
                    dd, mm, yyyy = parsed.split('/')
                    key = (mm, yyyy)
                    is_month_only = bool(mon_only_pat.match(date_match.strip()))
                    day_val = int(dd)
                    prev_day, _ = best.get(key, (0, None))
                    if not is_month_only and day_val > 1:
                        if prev_day <= 1:
                            best[key] = (day_val, parsed)
                    elif day_val > prev_day:
                        best[key] = (day_val, parsed)
        return {v for _, v in best.values()}

    def find_issuing_date(self, ocr_text, dob_str=None, expiry_str=None):
        candidates = self._collect_all_dates(ocr_text)
        if not candidates:
            return 'Not Found'
        dob = None
        if dob_str and dob_str != 'Not Found':
            try:
                dob = parser.parse(dob_str, dayfirst=True).date()
            except Exception:
                pass
        expiry = None
        if expiry_str and expiry_str != 'Not Found':
            try:
                expiry = parser.parse(expiry_str, dayfirst=True).date()
            except Exception:
                pass
        scored_candidates = []
        issue_keywords = ['ISSUE', 'ISSUANCE', 'ISSUED', 'DATE OF ISSUE', 'DATE OF ISSUANCE']
        ocr_blob = " ".join(ocr_text).upper()
        for cand_str in candidates:
            score = 0
            try:
                cand_date = parser.parse(cand_str, dayfirst=True).date()
            except Exception:
                continue
            for line in ocr_text:
                if str(cand_date.year) in line:
                    if any(kw in line.upper() for kw in issue_keywords):
                        score += 10
                        break
            if expiry:
                delta_years = expiry.year - cand_date.year
                if delta_years in [10, 5]:
                    score += 50
                elif delta_years > 0:
                    score += 10
                elif delta_years < 0:
                    score -= 100
            if dob and cand_date > dob:
                score += 5
            if expiry and cand_date < expiry:
                score += 5
            if dob_str == cand_str or expiry_str == cand_str:
                score -= 100
            if dob and cand_date == dob:
                score -= 100
            if expiry and cand_date == expiry:
                score -= 100
            scored_candidates.append((score, cand_str))
        if not scored_candidates:
            return 'Not Found'
        scored_candidates.sort(key=lambda x: (-x[0], x[1]))
        return scored_candidates[0][1]

    def find_authority(self, ocr_text):
        keywords = ['ISSUING AUTHORITY', 'ISSUED BY', 'AUTHORITY', 'ISSUING OFFICE', 'PLACE OF ISSUE']
        for line in ocr_text:
            for keyword in keywords:
                if keyword in line.upper():
                    authority = line.upper().split(keyword)[-1].strip()
                    return authority
        return 'Not Found'

    def _extract_labeled_fields(self, ocr_lines, label_map):
        normalized_lines = [self._normalize_ocr_line(line) for line in ocr_lines if line.strip()]
        results = {field: 'Not Found' for field in label_map.keys()}
        for i, line in enumerate(normalized_lines):
            upper_line = line.upper()
            for field, labels in label_map.items():
                if results[field] != 'Not Found':
                    continue
                for label in labels:
                    if label.upper() in upper_line:
                        if ':' in line:
                            value = line.split(':', 1)[1].strip()
                        else:
                            parts = re.split(re.escape(label), line, flags=re.IGNORECASE, maxsplit=1)
                            value = parts[1].strip() if len(parts) > 1 else ''
                        if not value and i + 1 < len(normalized_lines):
                            value = normalized_lines[i + 1].strip()
                        value = value.lstrip('.').strip()
                        results[field] = value if value else 'Not Found'
                        break
        return results

    def _extract_date_for_label(self, ocr_lines, label):
        normalized_lines = [self._normalize_ocr_line(line) for line in ocr_lines if line.strip()]
        label_upper = label.upper()
        date_patterns = self._get_all_date_patterns()
        for i, line in enumerate(normalized_lines):
            if label_upper in line.upper():
                for pattern in date_patterns:
                    for date_match in re.findall(pattern, line, re.IGNORECASE):
                        parsed_date = self._parse_ocr_date(date_match)
                        if parsed_date:
                            return parsed_date
                if i + 1 < len(normalized_lines):
                    next_line = normalized_lines[i + 1]
                    if not any(header in next_line.upper() for header in ['NAME', 'PASSPORT', 'AUTHORITY', 'BIRTH']):
                        for pattern in date_patterns:
                            for date_match in re.findall(pattern, next_line, re.IGNORECASE):
                                parsed_date = self._parse_ocr_date(date_match)
                                if parsed_date:
                                    return parsed_date
        return 'Not Found'

    def _extract_date_for_labels(self, ocr_lines, labels):
        for label in labels:
            found = self._extract_date_for_label(ocr_lines, label)
            if found != 'Not Found':
                return found
        return 'Not Found'

    def _extract_date_near_keywords(self, ocr_lines, keywords):
        normalized_lines = [self._normalize_ocr_line(line) for line in ocr_lines if line.strip()]
        keywords_upper = [k.upper() for k in keywords]
        date_patterns = self._get_all_date_patterns()
        month_pat = r'(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*'
        for i, line in enumerate(normalized_lines):
            upper_line = line.upper()
            if any(k in upper_line for k in keywords_upper):
                for pattern in date_patterns:
                    for m in re.finditer(pattern, line, re.IGNORECASE):
                        parsed = self._parse_ocr_date(m.group())
                        if parsed:
                            return parsed
                for offset in (1, 2):
                    if i + offset < len(normalized_lines):
                        next_line = normalized_lines[i + offset]
                        for pattern in date_patterns:
                            for m in re.finditer(pattern, next_line, re.IGNORECASE):
                                parsed = self._parse_ocr_date(m.group())
                                if parsed:
                                    return parsed
                block = ' '.join(normalized_lines[i:i + 6])
                for pattern in date_patterns:
                    for m in re.finditer(pattern, block, re.IGNORECASE):
                        parsed = self._parse_ocr_date(m.group())
                        if parsed:
                            return parsed
                pat = r'(\d{1,2})\s+' + month_pat + r'(?:\s+' + month_pat + r')?\s+(\d{2,4})'
                for m in re.finditer(pat, block, re.IGNORECASE):
                    parts = m.group(0).split()
                    s = f"{parts[0]} {parts[1]} {parts[-1]}"
                    parsed = self._parse_ocr_date(s)
                    if parsed:
                        return parsed
        return 'Not Found'

    def _split_date_components(self, date_string):
        if not date_string or date_string == 'Not Found':
            return 'Not Found', 'Not Found', 'Not Found'
        try:
            parsed = parser.parse(date_string, dayfirst=True).date()
        except (ValueError, TypeError):
            return 'Not Found', 'Not Found', 'Not Found'
        return f'{parsed.day:02d}', f'{parsed.month:02d}', f'{parsed.year:04d}'

    def _locate_roi_in_full_image(self, full_img, roi_img):
        if full_img is None or roi_img is None:
            return None
        if full_img.size == 0 or roi_img.size == 0:
            return None
        import numpy as np
        full_gray = cv2.cvtColor(full_img, cv2.COLOR_BGR2GRAY) if len(full_img.shape) == 3 else full_img
        roi_gray = cv2.cvtColor(roi_img, cv2.COLOR_BGR2GRAY) if len(roi_img.shape) == 3 else roi_img
        if full_gray.dtype != np.uint8:
            full_gray = (full_gray * 255).clip(0, 255).astype(np.uint8)
        if roi_gray.dtype != np.uint8:
            roi_gray = (roi_gray * 255).clip(0, 255).astype(np.uint8)
        fh, fw = full_gray.shape[:2]
        rh, rw = roi_gray.shape[:2]
        if rh >= fh or rw >= fw:
            return None
        res = cv2.matchTemplate(full_gray, roi_gray, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(res)
        if max_val < 0.55:
            return None
        x, y = max_loc
        return x, y, rw, rh
