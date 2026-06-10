---
name: add-country
description: Scaffold adding support for a new passport type — runs extraction, identifies failed fields, and guides targeted fixes in passport_data_extractor.py
disable-model-invocation: true
---

Usage: /add-country <CountryName> <image_path>

Steps:
1. Copy the test image into images/ with a descriptive name (e.g. Mexico.jpg)
2. Run extraction against the image:
   python -c "
   from passport_data_extractor import PassportDataExtractor
   p = PassportDataExtractor('desktop_app/data/country_codes.json', gpu=False)
   result = p.extract('<image_path>')
   import json; print(json.dumps(result, indent=2))
   "
3. Show which fields are missing or wrong (surname, given_names, dob, expiry, nationality, mrz)
4. Check passport_data_extractor.py for existing country-specific overrides (search for existing country name handling)
5. Suggest targeted regex or parsing fixes for the failed fields
6. After fixing, re-run extraction to confirm all fields are populated correctly
