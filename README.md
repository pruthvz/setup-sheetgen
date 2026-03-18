# SheetGen

SheetGen generates setup/installation sheets from an Excel part list and a Word template.

## What it does
- Reads rows from your Excel file.
- Maps the row data into fields in your Word template.
- Exports one output per row (DOCX, and optionally PDF).

## How to run
1. Install dependencies: `pip install -r requirements.txt`
2. Start the app: `python app.py`
3. In the UI, select your Excel file, Word template, and output folder, then generate sheets.

## License
This project is licensed under the Apache License, Version 2.0. See `LICENSE` for details.

# ⚙️ SheetGen
Generate setup sheets from Excel data and a Word template.

## What it does
SheetGen fills your Word template using row data from your Excel file, generating one output per row (DOCX and optionally PDF).

## Quick start
- Install: `pip install -r requirements.txt`
- Run: `python app.py`

<!--
---

## What it does
1. You point it at your **Excel file** and your **Word template** — once
2. The AI **automatically learns** the structure of both files
3. It parses every part code (like `ZZ9978318931PXX_100X2000_T1_EU_LH_F`) intelligently
4. Fills your Word template and saves each row as a **named PDF**

The AI analysis is cached — so the second run is instant. Tick **"Re-analyse files"** only if your template or Excel structure changes.

---

## Quick Start

### 1. Install Python
Download from https://python.org (3.10 or newer)

### 2. Install dependencies
Open a terminal / command prompt in this folder and run:
```
pip install -r requirements.txt
```

> **Note on docx2pdf**: This requires Microsoft Word to be installed on your machine (Windows or Mac). If you don't have Word, the tool will still work — it just won't produce PDFs, only .docx files.

### 3. Get an Anthropic API key
- Go to https://console.anthropic.com
- Create an account and generate an API key
- It costs fractions of a penny per run

### 4. Run the app
```
python app.py
```

### 5. First-time setup
- Click **⚙ Settings** and paste your API key
- Browse to your **Excel file**
- Browse to your **Word template**
- Choose an **output folder**
- Click **▶ GENERATE SETUP SHEETS**

The AI will analyse your files on first run (takes ~10 seconds), then generate all setup sheets automatically.

---

## How the AI learns your files

### Excel file
The AI reads the first 30 rows and figures out:
- Which column has part codes
- Which column has DNC/job numbers
- Any other relevant columns (quantity, material, date, etc.)

### Word template
The AI reads all text in the document and identifies:
- Every label/placeholder (like "Order No:", "Width:", "Machine:")
- How to map your part code data to those labels

This mapping is **saved automatically** so you don't repeat it every time.

---

## Part code parsing
Codes like `ZZ9978318931PXX_100X2000_T1_EU_LH_F` are parsed as:
| Part of code | Meaning |
|---|---|
| `ZZ9978318931` | Order number |
| `PXX` | Part suffix |
| `100X2000` | Width × Length (mm) |
| `T1` | Thickness |
| `EU` | Region |
| `LH` | Left-hand |
| `F` | Finish/door type |

The AI handles variations in format automatically.

---

## Troubleshooting

**"No API key found"** → Click Settings and add your Anthropic API key

**"No data rows found"** → Tick "Re-analyse files" and try again — the AI may have guessed the wrong column

**PDF not generated** → Make sure Microsoft Word is installed and `docx2pdf` is installed. The .docx file is still saved.

**Template not filling correctly** → Tick "Re-analyse files". You can also open `~/.setup_sheet_config.json` to see/edit the detected field mapping.

---

## Building the exe

To build a standalone `.exe` with PyInstaller:

```
pip install pyinstaller
pyinstaller SetupSheetGenerator.spec
```

The exe will be in `dist/SheetGen.exe`.

### OCR for scanned PDFs (Rename tool)

The Rename tool can read DNC numbers from **scanned** PDFs using OCR. For this to work:

**Option A — Bundle with exe (recommended for work PCs):**

1. Install Tesseract on your build machine: https://github.com/UB-Mannheim/tesseract/wiki  
2. Create a `tesseract_bundle` folder in the project root  
3. Copy the **entire** Tesseract folder (e.g. `C:\Program Files\Tesseract-OCR\`) into `tesseract_bundle/`:
   - `tesseract.exe`, all `.dll` files, and the `tessdata/` folder
   - Do **not** copy only tesseract.exe and tessdata — the exe needs the DLLs (e.g. `leptonica.dll`) to run
4. Rebuild the exe — the spec will bundle Tesseract automatically

**Option B — Run from Python without bundling:**

- Install Tesseract from the link above and add it to PATH, **or**
- Use `tesseract_bundle/` with the **full** Tesseract folder (including all DLLs) as above

If you skip this, the exe will still work for **digital** PDFs (text extraction). For scanned PDFs, OCR will fail without a working Tesseract.

---

## File locations
- Config & cache: `~/.setup_sheet_config.json` (your home folder)
- Output files: wherever you set the Output Folder in the app

---

-->

## License
Licensed under the Apache License, Version 2.0. See `LICENSE` for details.
