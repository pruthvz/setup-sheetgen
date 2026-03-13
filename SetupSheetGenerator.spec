# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
import os

datas = []
binaries = []

# Bundle Tesseract for OCR (scanned PDF rename). Create tesseract_bundle/ with:
#   tesseract_bundle/tesseract.exe
#   tesseract_bundle/tessdata/eng.traineddata (and other .traineddata as needed)
# Get these by installing Tesseract from https://github.com/UB-Mannheim/tesseract/wiki
# then copy from e.g. C:\Program Files\Tesseract-OCR\
TESSERACT_BUNDLE = os.path.join(SPECPATH, "tesseract_bundle")
if os.path.isdir(TESSERACT_BUNDLE):
    tess_exe = os.path.join(TESSERACT_BUNDLE, "tesseract.exe")
    tessdata_dir = os.path.join(TESSERACT_BUNDLE, "tessdata")
    if os.path.isfile(tess_exe) and os.path.isdir(tessdata_dir):
        datas.append((TESSERACT_BUNDLE, "tesseract"))
hiddenimports = ['anthropic', 'openai', 'openpyxl', 'docx', 'docx2pdf', 'pypdf', 'win32com', 'win32com.client', 'pywintypes',
                'fitz', 'pymupdf', 'pytesseract', 'PIL', 'PIL.Image', 'io']
tmp_ret = collect_all('anthropic')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('openai')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pymupdf')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='SheetGen',
    icon='sheetgen.ico',
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
)
