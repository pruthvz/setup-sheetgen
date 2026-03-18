# -*- mode: python ; coding: utf-8 -*-


from pathlib import Path

# PyInstaller executes this spec via `exec(...)` where `__file__` may be undefined.
# So we base paths on the current working directory (run pyinstaller from this folder).
root = Path.cwd()

# If you create a `tesseract_bundle/` folder (per README) containing:
#   tesseract.exe + tessdata/
# then pyinstaller will bundle it for OCR in the Rename tool.
tesseract_bundle_dir = root / "tesseract_bundle"
tesseract_datas = []
if tesseract_bundle_dir.exists():
    tesseract_datas = [(str(tesseract_bundle_dir), "tesseract")]

analysis_main = Analysis(
    ['redesign.py'],
    pathex=[],
    binaries=[],
    datas=tesseract_datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz_main = PYZ(analysis_main.pure)

exe_main = EXE(
    pyz_main,
    analysis_main.scripts,
    analysis_main.binaries,
    analysis_main.datas,
    [],
    name='SheetGen',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['sheetgen.ico'],
)

analysis_updater = Analysis(
    ['updater.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz_updater = PYZ(analysis_updater.pure)

exe_updater = EXE(
    pyz_updater,
    analysis_updater.scripts,
    analysis_updater.binaries,
    analysis_updater.datas,
    [],
    name='SheetGenUpdater',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[],
)
