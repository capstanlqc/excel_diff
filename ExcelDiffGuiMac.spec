# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

# Anchor to the current working directory when invoking PyInstaller
APP_DIR = Path.cwd()
ICON_ICNS = str(APP_DIR / "icons" / "icon.icns")

# Collect data files:
# - All JSONs from ./locales preserved under locales/
# - Entire icons directory preserved under icons/
locales_src = APP_DIR / "locales"
icons_src = APP_DIR / "icons"

datas = []
# Add all JSON files from locales into the bundle under "locales"
if locales_src.exists():
    for p in locales_src.iterdir():
        if p.is_file() and p.suffix.lower() == ".json":
            datas.append((str(p), "locales"))

# Add entire icons directory (contains PNGs and the .icns)
if icons_src.exists():
    datas.append((str(icons_src), "icons"))

a = Analysis(
    ['excel_diff_gui.py'],           # GUI is the entry-point; it can import/use excel_diff.py
    pathex=[str(APP_DIR)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,                       # ensures excel_diff_gui.py launches
    [],
    exclude_binaries=True,
    name='ExcelDiff',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                       # keep False on macOS; UPX disabled due to compatibility notes
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[ICON_ICNS],                # .icns for macOS app icon
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,                         # includes locales and icons
    strip=False,
    upx=False,
    upx_exclude=[],
    name='ExcelDiff',
)

app = BUNDLE(
    coll,
    name='ExcelDiff.app',
    icon=ICON_ICNS,
    bundle_identifier='be.capstan.exceldiff',
)
