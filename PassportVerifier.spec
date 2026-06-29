# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
from PyInstaller.utils.hooks import copy_metadata
from pathlib import Path

datas = [('data/country_codes.json', 'data')]
binaries = []
hiddenimports = []
datas += copy_metadata('imageio')
tmp_ret = collect_all('torch')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Bundle MRZScanner (DocsaidLab) with its ONNX models
tmp_ret = collect_all('mrzscanner')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Bundle DocAligner (DocsaidLab) with its ONNX model
tmp_ret = collect_all('docaligner')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Bundle capybara — shared DocsaidLab ONNX inference runtime
tmp_ret = collect_all('capybara')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Bundle fastmrz MRZ parser
tmp_ret = collect_all('fastmrz')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Bundle onnxruntime (needed by capybara/mrzscanner on Windows)
tmp_ret = collect_all('onnxruntime')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Bundle turbojpeg.dll (needed by capybara on Windows — libjpeg-turbo)
_turbojpeg_dll = 'C:/libjpeg-turbo64/bin/turbojpeg.dll'
if Path(_turbojpeg_dll).exists():
    binaries.append((_turbojpeg_dll, '.'))

# Bundle local Tesseract runtime (no system install needed on target machine).
vendor_tesseract = Path("vendor") / "tesseract"
if vendor_tesseract.exists():
    for src in vendor_tesseract.rglob("*"):
        if src.is_file():
            rel_parent = src.parent.relative_to(vendor_tesseract)
            dest = Path("tesseract") / rel_parent
            datas.append((str(src), str(dest)))


a = Analysis(
    ['desktop_app\\main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['rthooks/pyi_rth_turbojpeg.py'],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PassportVerifier',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PassportVerifier',
)
