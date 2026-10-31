# CSV to Excel Desktop Converter

A small, local desktop utility that converts one or more CSV files into a single
`.xlsx` workbook. Each input CSV is written to its own worksheet. The app uses
Tkinter for its single-window interface, pandas for CSV parsing, and openpyxl
for workbook creation. It does not upload files or require a browser.

## Features

- Select multiple CSV files in one operation.
- Automatically detects comma, semicolon, or tab delimiters.
- Detects common text encodings using `charset-normalizer`.
- Writes all parsed fields as literal Excel text. This helps preserve leading
  zeros, long IDs, date-like values, and strings beginning with `=`.
- Preserves duplicate header rows and blank records by treating the entire CSV
  as data; no header row or index is added or removed.
- Clear success and error dialogs.

## Requirements

- Python 3.10 or newer
- Tkinter (included with most standard Python installations; on some Linux
  distributions install the `python3-tk` package)
- Dependencies in `requirements.txt`

## Run from source

```bash
python -m venv .venv

# Windows
.venv\Scripts\\activate

# macOS/Linux
source .venv/bin/activate

python -m pip install -r requirements.txt
python csv_to_excel.py
```

Click **Select CSV**, choose one or more files, then click **Convert** and
choose where to save the workbook.

## Build a Windows executable

On a Windows machine, run:

```bat
build_windows.bat
```

The script installs the dependencies and PyInstaller, then creates a
single-folder application under `dist\CSVtoExcel\\`. Run
`dist\CSVtoExcel\\CSVtoExcel.exe`.

To build a single-file executable instead, change the PyInstaller command in
the batch file from `--onedir` to `--onefile`. Building on Windows is
recommended for a Windows executable; PyInstaller does not cross-compile
reliably between operating systems.

## Customize / extend

- **UI:** `ConverterApp` in `csv_to_excel.py` contains the Tkinter layout and
  button callbacks.
- **Encoding:** `detect_encoding()` uses `charset-normalizer`. Add an explicit
  fallback policy there if your environment has legacy files that cannot be
  detected reliably.
- **Delimiter detection:** `detect_delimiter()` supports comma, semicolon, and
  tab. Add other delimiters to its `csv.Sniffer` candidate list if needed.
- **Workbook output:** `write_workbook()` creates one worksheet per CSV.
  Modify this function to add formatting, filters, column widths, or a
  different sheet naming convention.

## Data fidelity notes

The converter intentionally imports values as **text**, rather than allowing
Excel to infer numbers, dates, or formulas. This avoids common silent changes
such as dropping leading zeros or converting long identifiers to scientific
notation. CSV quoting and escaping are interpreted according to CSV rules.

CSV is a text format and does not carry spreadsheet formulas, cell types,
formatting, or a distinction between every possible kind of empty cell.
Excel also imposes limits of 1,048,576 rows, 16,384 columns, and 32,767
characters per cell. Files exceeding those limits, or containing characters
that Excel cannot store, produce a clear error instead of a claimed success.

The app writes one workbook for the selected set of CSVs. If two input files
have the same base filename, worksheet names are made unique automatically.
