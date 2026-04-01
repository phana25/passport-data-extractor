#!/bin/bash
# MacOS release packaging script for Passport Data Extractor
# This script bundles the .app into a standard zip file for the auto-updater.

APP_NAME="PassportVerifier.app"
DIST_DIR="dist"

echo "📦 Packaging macOS application..."

if [ ! -d "$DIST_DIR/$APP_NAME" ]; then
    echo "❌ Error: $DIST_DIR/$APP_NAME not found. Please run 'pyinstaller --noconfirm Passport-Data-Extractor.spec' first."
    exit 1
fi

# Get version from __init__.py
VERSION=$(python3 -c "import sys; sys.path.append('.'); from desktop_app import __version__; print(__version__)")
ZIP_NAME="PassportVerifier_v${VERSION}_mac.zip"

echo "📁 Version detected: $VERSION"
echo "🤐 Creating zip: $DIST_DIR/$ZIP_NAME"

cd "$DIST_DIR"
zip -r -q "$ZIP_NAME" "$APP_NAME"
cd ..

echo "✅ Success! macOS bundle created at: $DIST_DIR/$ZIP_NAME"
