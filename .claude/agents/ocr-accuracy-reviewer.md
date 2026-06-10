---
name: ocr-accuracy-reviewer
description: Reviews changes to passport_data_extractor.py for regressions across country-specific edge cases and known failure modes
---

When reviewing changes to extraction logic:

1. Identify which fields and parsing paths were modified (MRZ, name, date, nationality, etc.)
2. Check against known tricky cases:
   - Names with 2 given names (recent fix in 3d58857)
   - Taiwan passport format (recent fix in 9b037c1)
   - Chinese/PRC passports
   - Passports with no MRZ (no_mrz.jpg fallback path)
   - Employee-style ID cards (employee.png, employee2.jpg)
3. Flag any regex change that could affect date parsing — dateutil and manual patterns coexist; verify both paths still work
4. Verify NAME_STOPWORDS still makes sense for the modified name extraction flow
5. Check if changes to `_extract_from_mrz` could break the `_extract_without_mrz` fallback or vice versa
6. Recommend running `python validate_error_images.py` if core MRZ parsing, name extraction, or date logic changed
