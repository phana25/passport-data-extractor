---
name: release
description: Guides through the full release checklist — version bump, changelog summary, version.json URL update, and GitHub release draft
disable-model-invocation: true
---

Steps:
1. Ask for the new version number (e.g. 1.1.5)
2. Update version.json with the new version and expected GitHub release URLs:
   - windows_url: https://github.com/phana25/passport-data-extractor/releases/download/v{VERSION}/Passport-Data-Extractor-Setup-v{VERSION}.exe
   - mac_url: https://github.com/phana25/passport-data-extractor/releases/download/v{VERSION}/PassportVerifier_v{VERSION}_mac.zip
3. Summarize git commits since the last tag as release notes
4. Create a GitHub release draft using: gh release create v{VERSION} --draft --title "v{VERSION}" --notes "{RELEASE_NOTES}"
5. Remind user to build and upload artifacts for both Windows and macOS before publishing
