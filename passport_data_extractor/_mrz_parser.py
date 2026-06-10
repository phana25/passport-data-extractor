import re
import string as st
import unicodedata
import cv2
import pytesseract


class MRZMixin:
    _NAME_STOPWORDS = {
        'ADMINISTRATION', 'IMMIGRATION', 'NATIONAL', 'REPUBLIC', 'PEOPLE',
        'CHINA', 'PRC', 'AUTHORITY', 'PASSPORT', 'MINISTRY', 'FOREIGN',
        'OFFICE', 'ISSUE', 'ISSUING', 'BEARER',
        'SURNAME', 'GSURNAME', 'GIVEN', 'NAMES', 'GIVENNAMES', 'SURAME',
    }

    def _normalize_mrz_name_token(self, token, is_given_name=False):
        value = (token or '').strip().upper()
        value = value.replace('«', '<').replace('|', '<')
        value = re.sub(r'<+', ' ', value)
        value = re.sub(r'[^A-Z ]+', ' ', value)
        value = re.sub(r'\s{2,}', ' ', value).strip()
        value = re.sub(r'\b([A-Z])\1{2,}\b$', '', value).strip()
        tokens = [t for t in value.split(' ') if t]

        while tokens and re.match(r'^([A-Z])\1{2,}$', tokens[-1]):
            tokens.pop()

        tokens = [re.sub(r'([A-Z])\1{2,}$', '', t) for t in tokens]
        tokens = [t for t in tokens if t]

        if is_given_name:
            expanded_tokens: list[str] = []
            for t in tokens:
                m = re.match(r'^([A-Z]{2,})K([A-Z]{2,})$', t)
                if m:
                    expanded_tokens.extend([m.group(1), m.group(2)])
                else:
                    expanded_tokens.append(t)
            tokens = expanded_tokens

        if any(len(t) > 1 for t in tokens):
            trailing_single_count = 0
            for t in reversed(tokens):
                if len(t) == 1:
                    trailing_single_count += 1
                else:
                    break
            if trailing_single_count >= 2:
                tokens = tokens[:len(tokens) - trailing_single_count]
            if tokens and tokens[-1] in ('K', 'X') and len(tokens) >= 2:
                if any(len(t) > 1 for t in tokens[:-1]):
                    tokens.pop()

        # Strip trailing tokens that are common OCR misreadings of `<` filler characters
        _FILLER = {
            'ES', 'KS', 'SK', 'EK', 'KE', 'CS', 'CK', 'KC', 'EX', 'XE',
            'FO', 'OF', 'FE', 'EF', 'CE', 'EC',
            'CSE', 'KEK', 'EKS', 'SKE', 'KES', 'CKE', 'ESK', 'SKS', 'EKE', 'CSK',
        }
        while len(tokens) > 1 and tokens[-1] in _FILLER:
            if any(len(t) > len(tokens[-1]) for t in tokens[:-1]):
                tokens.pop()
            else:
                break

        # Strip filler suffix concatenated directly onto the last token (e.g. "QINGBAOFO" → "QINGBAO")
        _CONCAT_FILLER = {'FO', 'OF', 'FE', 'EF', 'CE', 'EC', 'EK', 'KE', 'SK', 'KS', 'KC', 'CK'}
        if tokens:
            last = tokens[-1]
            for sfx in sorted(_CONCAT_FILLER, key=len, reverse=True):
                if last.endswith(sfx) and len(last) - len(sfx) >= 4:
                    tokens[-1] = last[:-len(sfx)]
                    break

        # Strip trailing garbage tokens: long tokens with very few unique letters
        # (e.g. "EDEIDESS" → 8 chars, 4 unique = OCR watermark noise)
        while len(tokens) > 1 and len(tokens[-1]) >= 7 and len(set(tokens[-1])) <= 4:
            tokens.pop()

        return ' '.join(tokens).strip()

    def _mrz_name_quality(self, given, surname):
        full = f"{surname} {given}".strip()
        if not full:
            return -1
        letters = len(re.findall(r'[A-Z]', full))
        tokens = [t for t in re.findall(r'[A-Z]+', full) if t]
        spaces = full.count(' ')
        penalties = 0
        if re.search(r'([A-Z])\1{3,}', full):
            penalties += 40
        elif re.search(r'([A-Z])\1{2,}', full):
            penalties += 15
        penalties += sum(1 for t in tokens if len(t) == 1) * 5
        if letters < 3:
            penalties += 10
        if any(len(t) > 18 for t in tokens):
            penalties += 40
        if 4 <= letters <= 25:
            penalties -= 2
        return letters + spaces - penalties

    def _is_suspicious_name(self, given, surname):
        full = f"{surname} {given}".strip().upper()
        if not full:
            return True
        # Raw word tokens (may contain digits); any digit in a name token = garbage
        raw_tokens = full.split()
        if any(any(c.isdigit() for c in t) for t in raw_tokens):
            return True
        tokens = [t for t in re.findall(r'[A-Z]+', full) if t]
        if not tokens:
            return True
        token_set = set(tokens)
        if ('SURNAME' in token_set or 'SURAME' in token_set) and ('GIVEN' in token_set or 'NAMES' in token_set):
            return True
        single_letter_tokens = [t for t in tokens if len(t) == 1]
        repeated_letter_blobs = [t for t in tokens if re.match(r'^([A-Z])\1{2,}$', t)]
        if repeated_letter_blobs:
            return True
        if any(len(set(t)) == 1 and len(t) >= 2 for t in tokens):
            return True
        if any(re.search(r'([A-Z])\1{2,}', t) for t in tokens):
            return True
        if len(single_letter_tokens) >= 2:
            return True
        if any(t in self._NAME_STOPWORDS for t in tokens):
            return True
        # Stopword concatenated with another word (e.g. "ISSUINGOFFICE" starts with "ISSUING")
        if any(any(t.startswith(sw) for sw in self._NAME_STOPWORDS) for t in tokens if len(t) > 4):
            return True
        if any(len(t) >= 18 for t in tokens):
            return True
        # Concatenated OCR artifact containing date-related words (e.g. "GNFININEXPIREDDATE")
        _DATE_EMBEDDED = frozenset(['EXPIRED', 'EXPIRY', 'EXPIRE'])
        if any(sub in t for t in tokens for sub in _DATE_EMBEDDED):
            return True
        # Token starting with "MRZ" = OCR watermark artifact from the "MRZ" label printed on the passport
        if any(t.startswith('MRZ') for t in tokens):
            return True
        # KK at start = OCR misread of two adjacent << filler chars
        if any(t.startswith('KK') for t in tokens):
            return True
        # Long token with very few unique letters = OCR garbage (e.g. "DSDDEEDEDDC")
        if any(len(t) >= 9 and len(set(t)) <= 4 for t in tokens):
            return True
        # Myanmar uses PVMMR (two-letter doc code PV + country MMR) — if it leaks into the name field it is garbage
        if any(t.startswith('PVMMR') or t.startswith('PVMM') for t in tokens):
            return True
        _FILLER_ENDS = {'SK', 'ES', 'EK', 'KS', 'CS', 'CK', 'CSE', 'KEK', 'EKS'}
        given_toks = [t for t in re.findall(r'[A-Z]+', (given or '').upper()) if t]
        if (len(given_toks) == 1 and len(given_toks[0]) > 8
                and given_toks[0][-2:] in _FILLER_ENDS or
                len(given_toks) == 1 and len(given_toks[0]) > 8
                and given_toks[0][-3:] in _FILLER_ENDS):
            return True
        return False

    def _parse_name_from_mrz_pair(self, line1, line2):
        """Parse surname and given name using fastmrz when available.

        Uses _parse_mrz with include_checkdigit=False so names are extracted
        even when check digits from OCR are unreliable.
        Falls back to _parse_name_from_mrz_line when fastmrz is unavailable
        or returns an empty result.
        """
        fast_mrz = getattr(self, '_fast_mrz', None)
        if fast_mrz is not None:
            try:
                # ICAO 2026 two-letter doc codes (e.g. PV for Myanmar, PP/PE for others) put a
                # letter in position 2. fastmrz expects the legacy P< format, so normalise
                # any letter subtype in position 2 to '<' before parsing.
                l1 = re.sub(r'^P([A-Z])([A-Z]{3})', lambda m: 'P<' + m.group(2), line1)
                # fastmrz requires exactly 44 chars per TD3 line — pad or trim
                l1 = (l1 + '<' * 44)[:44]
                l2 = (line2 + '<' * 44)[:44]
                result = fast_mrz._parse_mrz(l1 + '\n' + l2, include_checkdigit=False)
                surname = (result.get('surname') or '').strip()
                given = (result.get('given_name') or '').strip()
                if surname or given:
                    return given, surname
            except Exception:
                pass
        return self._parse_name_from_mrz_line(line1)

    def _parse_name_from_mrz_line(self, line_a):
        if not line_a:
            return '', ''
        raw = re.sub(r'\s+', '', line_a).upper()
        raw = raw.replace('«', '<').replace('|', '<')
        # OCR commonly misreads the MRZ fill char `<` as `K` when it immediately follows
        # certain letters (e.g. FAUZIAH< → FAUZIAHK, EMILY< → EMILYK, RISKY< → RISKYK).
        # This substitution fixes `K<<` boundaries: letter+K+fill → letter+fill.
        raw = re.sub(r'([AEIOUYH])K(<<+)', r'\1\2', raw)
        if '<<' in raw:
            left, right = raw.split('<<', 1)
        else:
            return '', ''
        # Strip the 5-char MRZ header (doc-type + subtype + 3-letter country code).
        # Position 2 can be '<' (legacy single-letter) or a letter subtype for any country
        # using the ICAO 2026 two-letter doc code (PV for Myanmar, PP/PE/PD … for others).
        # two_letter_subtype flags any two-letter code: these passports pack all name parts
        # into the surname field with no '<<' separator, so the split must be done manually.
        two_letter_subtype = False
        if re.match(r'^P[A-Z<][A-Z]{3}', left):
            two_letter_subtype = bool(re.match(r'^P[A-Z][A-Z]{3}', left))
            left = left[5:]
        surname = self._normalize_mrz_name_token(left, is_given_name=False)
        given = self._normalize_mrz_name_token(right, is_given_name=True)

        # Fallback: if given is empty the << was filler at the end of the line,
        # meaning OCR misread the surname/given separator << as <K.
        # e.g. "MAI<KVAN<DUY<<<" → surname="MAI KVAN DUY", given=""
        # should be  → surname="MAI", given="VAN DUY"
        # Do NOT apply this for two-letter subtype passports: after stripping the
        # country prefix, every <K is a genuine name token starting with K
        # (e.g. CHI<YINN<KYIN), not an OCR misread of the << separator.
        if not given and surname and not two_letter_subtype:
            m = re.search(r'<K([A-Z][A-Z<]*)', left)
            if m:
                alt_sur = self._normalize_mrz_name_token(left[:m.start()], is_given_name=False)
                alt_giv = self._normalize_mrz_name_token(m.group(1), is_given_name=True)
                # Only accept the <K split when the left portion before <K contains
                # no further single-< separators — i.e. it looks like a clean single
                # surname (e.g. "MAI"), not a multi-word name where K begins a real token
                # (e.g. "SISKA<KURNIA" → don't split KURNIA into K+URNIA).
                left_before = left[:m.start()]
                if (alt_sur and alt_giv
                        and not self._is_suspicious_name(alt_giv, alt_sur)
                        and '<' not in left_before):
                    surname, given = alt_sur, alt_giv

        # Two-letter subtype passports pack all name parts into the surname field (no << separator).
        # Apply the last-word-is-surname split so the candidate enters scoring with
        # both given and surname filled — earning the complete-pair quality bonus.
        if two_letter_subtype and not given and surname and ' ' in surname:
            parts = surname.split()
            surname = parts[-1]
            given = ' '.join(parts[:-1])

        return given, surname

    def _extract_visual_latin_name(self, ocr_lines):
        stopwords = {
            'PEOPLE', 'REPUBLIC', 'CHINA', 'PASSPORT', 'MINISTRY', 'FOREIGN',
            'AFFAIRS', 'REQUESTS', 'AUTHORITIES', 'ALLOW', 'BEARER', 'ASSISTANCE',
            'DATE', 'BIRTH', 'ISSUE', 'EXPIRY', 'NATIONALITY', 'AUTHORITY',
            'SURNAME', 'GSURNAME', 'GIVEN', 'NAMES', 'GIVENNAMES', 'SURAME',
            # Myanmar Ministry of Home Affairs authority abbreviation — never a person name
            'MOHA',
        }
        _AUTHORITY_MARKERS = frozenset([
            'AUTHORITY', 'AUTHONTY', 'ISSUING', 'ISSUED', 'OFFICE', 'MINISTER',
            'MINISTRY', 'IMMIGRATION', 'MOHA',
        ])
        normalized = [self._normalize_ocr_line(line).upper() for line in ocr_lines]
        for idx, n in enumerate(normalized):
            n = n.replace('，', ',')
            m = re.search(r'\b([A-Z]{2,})\s*,\s*([A-Z][A-Z ]{1,})\b', n)
            if not m:
                continue
            # Skip if any of the 4 preceding lines mention an authority/issuing keyword —
            # this catches Myanmar "AUTHORITY: MOHA, YANGON" patterns.
            context_lines = normalized[max(0, idx - 4):idx]
            if any(marker in ctx for ctx in context_lines for marker in _AUTHORITY_MARKERS):
                continue
            # Also skip if the matched line itself contains an authority keyword outside the name.
            pre_match = n[:m.start()].split()
            if any(marker in pre_match for marker in _AUTHORITY_MARKERS):
                continue
            surname = self._normalize_mrz_name_token(m.group(1))
            given = self._normalize_mrz_name_token(m.group(2))
            if not surname or not given:
                continue
            surname_tokens = surname.split()
            given_tokens = given.split()
            all_tokens = surname_tokens + given_tokens
            if len(surname_tokens) != 1:
                continue
            if len(given_tokens) > 2:
                continue
            if any(len(t) == 1 for t in all_tokens):
                continue
            if any(t in stopwords for t in all_tokens):
                continue
            if len(''.join(all_tokens)) > 24:
                continue
            if not all(re.search(r'[AEIOU]', t) for t in all_tokens):
                continue
            return given, surname
        return '', ''

    @staticmethod
    def _strip_diacritics(text):
        return ''.join(
            c for c in unicodedata.normalize('NFD', text)
            if unicodedata.category(c) != 'Mn'
        )

    def _extract_given_by_surname_from_visual(self, ocr_lines, surname):
        s = self._normalize_mrz_name_token(surname)
        if not s:
            return ''
        for line in ocr_lines:
            n = self._strip_diacritics(self._normalize_ocr_line(line)).upper()
            n = re.sub(r'[^A-Z,.\- ]+', ' ', n)
            n = re.sub(r'[,.\-]+', ' ', n)
            tokens = [t for t in n.split() if t]
            if not tokens:
                continue
            for i, tok in enumerate(tokens):
                if tok != s:
                    continue
                following_chunks = []
                idx = i + 1
                _VISUAL_FILLER = {
                    'FO', 'OF', 'FE', 'EF', 'CE', 'EC', 'EK', 'KE', 'SK', 'KS',
                    'KC', 'CK', 'ES', 'KX', 'XK', 'EX', 'XE',
                    'EY', 'YE', 'YK', 'KY',  # MRZ fill-char OCR artifacts
                }
                while idx < len(tokens) and len(following_chunks) < 5:
                    nxt = tokens[idx]
                    idx += 1
                    candidate = self._normalize_mrz_name_token(nxt)
                    if not candidate:
                        continue
                    if candidate in self._NAME_STOPWORDS:
                        break
                    if len(candidate) > 14:
                        break
                    # Stop on filler artifacts regardless of how many chunks collected.
                    if len(candidate) <= 2 and candidate in _VISUAL_FILLER:
                        break
                    if len(candidate) == 1:
                        if candidate in ('M', 'F') or not following_chunks:
                            break
                        if idx < len(tokens):
                            next_candidate = self._normalize_mrz_name_token(tokens[idx])
                            if next_candidate and len(next_candidate) == 1 and next_candidate not in ('M', 'F'):
                                idx += 1
                                following_chunks.append(candidate + next_candidate)
                                continue
                        following_chunks.append(candidate)
                        continue
                    following_chunks.append(candidate)
                if following_chunks:
                    return ' '.join(following_chunks).strip()
        return ''

    def _extract_mrz_name_from_ocr_lines(self, ocr_lines):
        for line in ocr_lines:
            n = re.sub(r'\s+', '', (line or '').upper())
            n = n.replace('«', '<').replace('|', '<')
            if not re.match(r'^P[<O][A-Z]{3}[A-Z<]{8,}$', n):
                continue
            given, surname = self._parse_name_from_mrz_line(n)
            if given and surname and not self._is_suspicious_name(given, surname):
                return given, surname
        return '', ''

    def _extract_name_from_visual_labels(self, ocr_lines):
        SURNAME_LABELS = [
            'SURNAME', 'FAMILY NAME', 'LAST NAME', 'NOM', 'APELLIDOS', 'COGNOME',
            'FAMILIENNAME', 'ACHTERNAAM',
        ]
        GIVEN_LABELS = [
            'GIVEN NAMES', 'GIVEN NAME', 'FIRST NAME', 'PRENOM', 'PRENOMS',
            'NOMBRES', 'NOME', 'VORNAME', 'VOORNAMEN',
        ]
        normalized = [self._normalize_ocr_line(l).upper() for l in ocr_lines if l.strip()]
        surname = ''
        given = ''
        for i, line in enumerate(normalized):
            if not surname:
                for label in SURNAME_LABELS:
                    if label in line:
                        after = re.split(re.escape(label), line, maxsplit=1)[1]
                        after = re.sub(r'^[\s:./|,]+', '', after).strip()
                        val = self._normalize_mrz_name_token(after) if len(after) >= 2 else ''
                        if not val and i + 1 < len(normalized):
                            val = self._normalize_mrz_name_token(normalized[i + 1])
                        if val and not self._is_suspicious_name(val, ''):
                            surname = val
                        break
            if not given:
                for label in GIVEN_LABELS:
                    if label in line:
                        after = re.split(re.escape(label), line, maxsplit=1)[1]
                        after = re.sub(r'^[\s:./|,]+', '', after).strip()
                        val = self._normalize_mrz_name_token(after) if len(after) >= 2 else ''
                        if not val and i + 1 < len(normalized):
                            val = self._normalize_mrz_name_token(normalized[i + 1])
                        if val and not self._is_suspicious_name(val, ''):
                            given = val
                        break
            if surname and given:
                break

        # Fallback: standalone "Name:" label (employment card format, e.g. "Name AUNG MYO WIN")
        if not surname and not given:
            _non_name_tokens = {
                'EMPLOYED', 'SELF', 'COMPANY', 'ENTERPRISE', 'POSITION', 'BUSINESS',
                'CONSULTANT', 'CONSUKANT', 'CARD', 'MALE', 'FEMALE', 'NATIONALITY',
                'EMPLOYMENT', 'FOREIGN', 'REPUBLIC', 'MINISTRY', 'IMMIGRATION',
            }
            for i, line in enumerate(normalized):
                # Match standalone NAME — not inside SURNAME, GIVEN NAME, FAMILY NAME etc.
                if not re.search(r'(?<![A-Z])NAME(?![A-Z])', line):
                    continue
                after = re.split(r'(?<![A-Z])NAME(?![A-Z])', line, maxsplit=1)[1]
                after = re.sub(r'^[\s:./|,]+', '', after).strip()
                val = self._normalize_mrz_name_token(after) if len(after) >= 4 else ''
                if not val and i + 1 < len(normalized):
                    val = self._normalize_mrz_name_token(normalized[i + 1])
                if not val:
                    continue
                val_tokens = val.split()
                if (len(val_tokens) >= 2
                        and all(t.isalpha() for t in val_tokens)
                        and len(val_tokens[-1]) >= 3  # surname token must be substantial
                        and not any(t in _non_name_tokens for t in val_tokens)):
                    # last word = surname, rest = given (consistent with _split_name convention)
                    surname = val_tokens[-1]
                    given = ' '.join(val_tokens[:-1])
                    break

        return given, surname

    def _mrz_char_value(self, ch):
        if '0' <= ch <= '9':
            return ord(ch) - ord('0')
        if 'A' <= ch <= 'Z':
            return ord(ch) - ord('A') + 10
        return 0

    def _mrz_check_digit(self, data):
        weights = [7, 3, 1]
        total = sum(self._mrz_char_value(ch) * weights[i % 3] for i, ch in enumerate(data))
        return str(total % 10)

    def _preprocess_mrz_variants(self, img):
        if img is None or (hasattr(img, 'size') and img.size == 0):
            return []
        if len(img.shape) == 2:
            img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
        h, w = img.shape[:2]
        if h == 0 or w == 0:
            return []
        variants = []
        for target_h in (140, 200):
            scale = target_h / max(h, 1)
            target_w = max(int(w * scale), 900)
            resized = cv2.resize(img, (target_w, target_h), interpolation=cv2.INTER_LANCZOS4)
            gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
            _, thresh_otsu = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            variants.append(cv2.cvtColor(thresh_otsu, cv2.COLOR_GRAY2BGR))
            clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 4))
            enhanced = clahe.apply(gray)
            _, thresh_clahe = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            variants.append(cv2.cvtColor(thresh_clahe, cv2.COLOR_GRAY2BGR))
            variants.append(cv2.cvtColor(cv2.bitwise_not(thresh_otsu), cv2.COLOR_GRAY2BGR))
        variants.append(cv2.resize(img, (1110, 140)))
        return variants

    def _ocr_mrz_roi(self, roi_img):
        allowlist = st.ascii_letters + st.digits + '< '
        seen: set = set()
        all_lines: list = []

        def _add(line):
            key = re.sub(r'\s+', '', (line or '').upper()).replace('«', '<').replace('|', '<')
            if len(key) >= 5 and key not in seen:
                seen.add(key)
                all_lines.append(line)

        for variant in self._preprocess_mrz_variants(roi_img):
            for line in self.reader.readtext(variant, paragraph=False, detail=0, allowlist=allowlist):
                _add(line)

        if self._tesseract_available:
            tess_cfg = '--psm 6 --oem 3 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789< '
            for variant in self._preprocess_mrz_variants(roi_img)[:3]:
                try:
                    gray = cv2.cvtColor(variant, cv2.COLOR_BGR2GRAY) if len(variant.shape) == 3 else variant
                    text = pytesseract.image_to_string(gray, config=tess_cfg)
                    for line in text.splitlines():
                        _add(line.strip())
                except Exception:
                    pass

        return all_lines

    def _reconstruct_mrz_lines(self, ocr_lines):
        normalized = []
        for line in ocr_lines:
            n = re.sub(r'\s+', '', (line or '').upper()).replace('«', '<').replace('|', '<')
            n = re.sub(r'[^A-Z0-9<]', '', n)
            if len(n) >= 5:
                normalized.append(n)
        seen: set = set(normalized)
        result = list(normalized)
        for window in range(2, 5):
            for i in range(len(normalized) - window + 1):
                combined = ''.join(normalized[i:i + window])
                if 38 <= len(combined) <= 50:
                    candidate = (combined + '<' * 44)[:44]
                    if candidate not in seen:
                        seen.add(candidate)
                        result.append(candidate)
        return result

    def _correct_mrz_field_with_check(self, field, expected_check, numeric_only=False):
        if self._mrz_check_digit(field) == expected_check:
            return field
        DIGIT_SUBS = {
            'O': '0', 'Q': '0', 'D': '0', 'I': '1', 'L': '1', 'J': '1',
            'B': '8', 'S': '5', 'Z': '2', 'G': '6', 'T': '7',
        }
        ALPHA_SUBS = {'0': 'O', '1': 'I', '8': 'B', '5': 'S', '2': 'Z'}
        obvious = ''.join(
            DIGIT_SUBS.get(ch, ch) if (numeric_only or not ch.isdigit()) else ch
            for ch in field
        )
        if self._mrz_check_digit(obvious) == expected_check:
            return obvious
        subs = {**DIGIT_SUBS, **(ALPHA_SUBS if not numeric_only else {})}
        candidate = list(obvious)
        for i, ch in enumerate(candidate):
            if ch in subs:
                trial = candidate.copy()
                trial[i] = subs[ch]
                if self._mrz_check_digit(''.join(trial)) == expected_check:
                    return ''.join(trial)
        if len(field) <= 9:
            positions = [i for i, ch in enumerate(candidate) if ch in subs]
            for pi in range(len(positions)):
                for pj in range(pi + 1, len(positions)):
                    a, b_pos = positions[pi], positions[pj]
                    trial = candidate.copy()
                    trial[a] = subs[candidate[a]]
                    trial[b_pos] = subs[candidate[b_pos]]
                    if self._mrz_check_digit(''.join(trial)) == expected_check:
                        return ''.join(trial)
        return field

    def _auto_correct_mrz_line2(self, line2):
        if len(line2) != 44:
            return line2
        chars = list(line2)
        for start, end, chk_pos, num_only in [
            (0, 9, 9, False),
            (13, 19, 19, True),
            (21, 27, 27, True),
            (28, 42, 42, False),
        ]:
            raw = ''.join(chars[start:end])
            corrected = self._correct_mrz_field_with_check(raw, chars[chk_pos], num_only)
            for i, ch in enumerate(corrected):
                chars[start + i] = ch
        digit_to_alpha = {'0': 'O', '1': 'I', '8': 'B', '5': 'S', '2': 'Z'}
        for pos in range(10, 13):
            if chars[pos] in digit_to_alpha:
                chars[pos] = digit_to_alpha[chars[pos]]
        return ''.join(chars)

    def _normalize_mrz_line(self, line, line_index):
        s = re.sub(r'\s+', '', (line or '').upper())
        s = s.replace('«', '<').replace('|', '<')
        s = re.sub(r'[^A-Z0-9<]', '<', s)
        if line_index == 0:
            pass
        else:
            DIGIT_SUBS = {'O': '0', 'Q': '0', 'I': '1', 'L': '1', 'B': '8'}
            if len(s) == 44:
                chars = list(s)
                for rng in [(9, 10), (13, 20), (21, 28), (42, 44)]:
                    for i in range(*rng):
                        if chars[i] in DIGIT_SUBS:
                            chars[i] = DIGIT_SUBS[chars[i]]
                for i in range(0, 9):
                    if chars[i] in DIGIT_SUBS:
                        chars[i] = DIGIT_SUBS[chars[i]]
                s = ''.join(chars)
            elif len(s) == 30:
                chars = list(s)
                for rng in [(0, 7), (8, 15)]:
                    for i in range(*rng):
                        if chars[i] in DIGIT_SUBS:
                            chars[i] = DIGIT_SUBS[chars[i]]
                for i in range(0, 5):
                    if chars[i] in DIGIT_SUBS:
                        chars[i] = DIGIT_SUBS[chars[i]]
                s = ''.join(chars)
            else:
                s = s.replace('O', '0').replace('Q', '0')
                s = s.replace('I', '1').replace('L', '1')
                s = s.replace('B', '8')
        return s

    def _mrz_td1_validation_score(self, line1, line2, line3):
        score = 0
        details = {}
        if len(line1) != 30 or len(line2) != 30 or len(line3) != 30:
            return -100, {'length_ok': False}
        details['length_ok'] = True
        if line1[0] in ('I', 'A', 'C', 'V'):
            score += 8
        for key, data, chk in [
            ('doc_number', line1[5:14], line1[14]),
            ('dob', line2[0:6], line2[6]),
            ('expiry', line2[8:14], line2[14]),
        ]:
            ok = self._mrz_check_digit(data) == chk
            details[f'{key}_ok'] = ok
            if ok:
                score += 20
        composite_data = line1[5:30] + line2[0:7] + line2[8:15] + line2[18:29]
        composite_ok = self._mrz_check_digit(composite_data) == line2[29]
        details['composite_ok'] = composite_ok
        if composite_ok:
            score += 30
        details['check_pass_count'] = sum(1 for v in details.values() if v is True)
        return score, details

    def _auto_correct_mrz_td1(self, line1, line2):
        digit_to_alpha = {'0': 'O', '1': 'I', '8': 'B', '5': 'S', '2': 'Z'}
        if len(line1) == 30:
            chars1 = list(line1)
            corrected = self._correct_mrz_field_with_check(''.join(chars1[5:14]), chars1[14], False)
            for i, ch in enumerate(corrected):
                chars1[5 + i] = ch
            for pos in range(2, 5):
                if chars1[pos] in digit_to_alpha:
                    chars1[pos] = digit_to_alpha[chars1[pos]]
            line1 = ''.join(chars1)
        if len(line2) == 30:
            chars2 = list(line2)
            for start, end, chk_pos, num_only in [(0, 6, 6, True), (8, 14, 14, True)]:
                corrected = self._correct_mrz_field_with_check(
                    ''.join(chars2[start:end]), chars2[chk_pos], num_only)
                for i, ch in enumerate(corrected):
                    chars2[start + i] = ch
            for pos in range(15, 18):
                if chars2[pos] in digit_to_alpha:
                    chars2[pos] = digit_to_alpha[chars2[pos]]
            line2 = ''.join(chars2)
        return line1, line2

    def _mrz_validation_score(self, line1, line2):
        score = 0
        details = {}
        if len(line1) != 44 or len(line2) != 44:
            return -100, {'length_ok': False}
        details['length_ok'] = True
        if line1.startswith('P<'):
            score += 8
            details['doc_prefix_ok'] = True
        else:
            details['doc_prefix_ok'] = False
        checks = [
            ('passport', line2[0:9], line2[9]),
            ('dob', line2[13:19], line2[19]),
            ('expiry', line2[21:27], line2[27]),
            ('personal', line2[28:42], line2[42]),
        ]
        pass_count = 0
        for key, data, chk in checks:
            ok = self._mrz_check_digit(data) == chk
            details[f'{key}_ok'] = ok
            if ok:
                pass_count += 1
                score += 20
        composite_data = line2[0:10] + line2[13:20] + line2[21:43]
        composite_ok = self._mrz_check_digit(composite_data) == line2[43]
        details['composite_ok'] = composite_ok
        if composite_ok:
            pass_count += 1
            score += 30
        details['check_pass_count'] = pass_count
        # Country code consistency: for passport lines, the 3-letter code in line1[2:5]
        # should match the nationality in line2[10:13]. A mismatch (e.g. "KID" vs "IDN")
        # strongly indicates OCR garbling in the document-type/country prefix area.
        if line1[0] == 'P' and len(line1) >= 5 and len(line2) >= 13:
            l1_country = line1[2:5]
            l2_nat = line2[10:13]
            if (re.match(r'^[A-Z]{3}$', l1_country) and re.match(r'^[A-Z]{3}$', l2_nat)):
                if l1_country == l2_nat:
                    score += 6
                else:
                    score -= 6
        return score, details

    def _build_mrz_candidates(self, ocr_code_lines):
        all_fragments = self._reconstruct_mrz_lines(ocr_code_lines)
        seen_pairs: set = set()
        candidates = []
        for frag_a in all_fragments:
            for frag_b in all_fragments:
                if frag_a == frag_b:
                    continue
                l1 = self._normalize_mrz_line(frag_a, 0)
                l2 = self._normalize_mrz_line(frag_b, 1)
                if len(l1) < 30 or len(l2) < 30:
                    continue
                l1 = (l1 + '<' * 44)[:44]
                l2 = (l2 + '<' * 44)[:44]
                key = (l1, l2)
                if key in seen_pairs:
                    continue
                seen_pairs.add(key)
                candidates.append((l1, l2))
        return candidates

    def _select_best_mrz_candidate(self, ocr_code_lines, mrz_obj=None):
        candidates = self._build_mrz_candidates(ocr_code_lines)
        if not candidates:
            return None
        best = None
        for line1, line2 in candidates:
            line2 = self._auto_correct_mrz_line2(line2)
            base_score, details = self._mrz_validation_score(line1, line2)
            given, surname = self._parse_name_from_mrz_pair(line1, line2)
            name_score = self._mrz_name_quality(given, surname)
            suspicious_penalty = 35 if self._is_suspicious_name(given, surname) else 0
            total = base_score + name_score - suspicious_penalty
            rec = {
                'line1': line1, 'line2': line2,
                'given': given, 'surname': surname,
                'score': total, 'details': details,
            }
            if best is None or rec['score'] > best['score']:
                best = rec
        if best and mrz_obj:
            mrz_data = mrz_obj.to_dict()
            pg = self._normalize_mrz_name_token(mrz_data.get('names', ''), is_given_name=True)
            ps = self._normalize_mrz_name_token(mrz_data.get('surname', ''), is_given_name=False)
            if pg and ps and not self._is_suspicious_name(pg, ps):
                if self._mrz_name_quality(pg, ps) >= self._mrz_name_quality(best['given'], best['surname']):
                    best['given'], best['surname'] = pg, ps
        return best

    def _select_best_td1_candidate(self, ocr_lines):
        fragments = self._reconstruct_mrz_lines(ocr_lines)
        td1_frags = []
        for f in fragments:
            if 26 <= len(f) <= 34:
                n = self._normalize_mrz_line(f, 0)
                n = (n + '<' * 30)[:30]
                if n not in td1_frags:
                    td1_frags.append(n)
        if len(td1_frags) < 3:
            return None
        seen: set = set()
        best = None
        for i, l1 in enumerate(td1_frags):
            for j, l2_raw in enumerate(td1_frags):
                for k, l3 in enumerate(td1_frags):
                    if i == j or j == k or i == k:
                        continue
                    l2 = self._normalize_mrz_line(l2_raw, 1)
                    l2 = (l2 + '<' * 30)[:30]
                    key = (l1, l2, l3)
                    if key in seen:
                        continue
                    seen.add(key)
                    l1c, l2c = self._auto_correct_mrz_td1(l1, l2)
                    base_score, details = self._mrz_td1_validation_score(l1c, l2c, l3)
                    given, surname = self._parse_name_from_mrz_line(l3)
                    name_score = self._mrz_name_quality(given, surname)
                    suspicious_penalty = 35 if self._is_suspicious_name(given, surname) else 0
                    total = base_score + name_score - suspicious_penalty
                    rec = {
                        'line1': l1c, 'line2': l2c, 'line3': l3,
                        'given': given, 'surname': surname,
                        'score': total, 'details': details, 'mrz_type': 'TD1',
                    }
                    if best is None or total > best['score']:
                        best = rec
        return best
