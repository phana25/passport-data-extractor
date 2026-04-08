# Build and Release Guide: Passport Data Extractor

This guide explains how to generate production builds for both **Windows** and **macOS**.

---

## 📋 Pre-Build Checklist
Before building, ensure you have updated the version numbers and verified the update URLs:

1. **Check Version**: Ensure `__version__` in [__init__.py](file:///Users/sophana/Projects/Passport-Data-Extractor/desktop_app/__init__.py) is set correctly (e.g., `"1.1.0"`).
2. **Verify URLs**: check [updater.py](file:///Users/sophana/Projects/Passport-Data-Extractor/desktop_app/services/updater.py) to make sure `UPDATE_JSON_URL` and `FALLBACK_UPDATE_URL` are correct.
3. **Public Repo**: If you are about to release, ensure your GitHub repository or Gist is **Public**.

---

## 🪟 Windows Build Instructions
Windows builds use **PyInstaller** to create the application files and **Inno Setup** to create the `.exe` installer.

### Prerequisites (Windows)
- [Inno Setup 6](https://jrsoftware.org/isdl.php) installed.
- Microsoft VC++ Redistributable (`vc_redist.x64.exe`) placed in `third_party/`.

### Steps:
1. **Generate App Files**:
   Run this in your terminal (PowerShell or Command Prompt):
   ```powershell
   pyinstaller --noconfirm "Passport-Data-Extractor.spec"
   ```
2. **Build Installer**:
   Run the automated build script:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\build_installer.ps1
   ```
3. **Output**: Your installer (`Passport-Data-Extractor-Setup-v1.1.0.exe`) will be in the `installer_output/` folder (the version is read automatically from `desktop_app/__init__.py`).

---

## 🍎 macOS Build Instructions
macOS builds use **PyInstaller** to create a native `.app` bundle. Since you are currently on macOS, you can run these steps directly.

### Prerequisites (macOS)
- Tesseract installed via Homebrew: `brew install tesseract`.

### Steps:
1. **Generate .app Bundle**:
   Run this in your terminal:
   ```bash
   pyinstaller --noconfirm "Passport-Data-Extractor.spec"
   ```
2. **Verify**: Your application (`PassportVerifier.app`) will be created in the `dist/` folder.
3. **Create Distribution Zip**:
   Right-click `PassportVerifier.app` in Finder and select **Compress**, or run:
   ```bash
   cd dist && zip -r PassportVerifier_v1.1.0_mac.zip PassportVerifier.app
   ```
4. **Output**: The `.zip` file is what you should upload to your GitHub Release.

---

## 🚀 Releasing to GitHub (Step-by-Step)

Once you have your Windows `.exe` and macOS `.zip` files ready, follow these exact steps to push the update to your users:

### 1. Create the GitHub Release
1. Go to your repository on [GitHub](https://github.com/phana25/passport-data-extractor).
2. Click on **Releases** (on the right-side sidebar) → **Draft a new release**.
3. Click **Choose a tag** and type `v1.1.0` (this must match the version in your code).
4. For the **Release title**, use `v1.1.0 - New OCR and UI features`.
5. Under the "Attach binaries" area, drag and drop:
   - `Passport-Data-Extractor-Setup-v1.1.0.exe` (Windows)
   - `PassportVerifier_v1.1.0_mac.zip` (macOS)
6. Click **Publish release**.

### 2. Get the Direct Download Links
1. After publishing, you will see your files listed under the **Assets** section of the release.
2. **Right-click** on each file and select **Copy Link Address**.
   - Example Windows Link: `https://github.com/phana25/passport-data-extractor/releases/download/v1.1.0/Passport-Data-Extractor-Setup-v1.1.0.exe`
   - Example macOS Link: `https://github.com/phana25/passport-data-extractor/releases/download/v1.1.0/PassportVerifier_v1.1.0_mac.zip`

### 3. Update the Version Gist
1. Go to your [Public Gist](https://gist.github.com/phana25/14514d5de8ded8943899638598638b93).
2. Click **Edit**.
3. Update the `latest_version` to `"1.1.0"`.
4. Paste the links you copied in Step 2 into `windows_url` and `mac_url`.
5. Update the `release_notes` description if you'd like.
6. Click **Update public gist**.

### 🛠️ What happens next?
1. Any user running version **1.0.0** of your app will automatically see a **blue update banner** at the top.
2. When they click **Download**, the app will pull the files from the links you just put in the Gist.
3. The app will update itself and restart with your new **1.0.1** version!

