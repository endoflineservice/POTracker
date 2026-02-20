# POtrol

A lightweight app for fast purchase order entry and searching, while storing everything in Excel.

## What It Does

- Uses your existing workbook (`IT POs.xlsx`) as the main data store
- Lets you enter records in a fast form inside a Streamlit app
- Lets you search and view all entries in a table
- Includes a `Reason for Purchase` field in entry and search
- Creates a timestamped backup copy of the workbook every time you save
- Keeps only the most recent backup files (configurable)
- Uses custom POtrol branding assets in `assets/` (badge logo + app icon)

## Setup (Windows PowerShell)

```powershell
cd "c:\Users\cholt\OneDrive - Champagne Metals\Documents\GitHub\POTracker"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

```powershell
streamlit run potrol.py
```

When the app opens:

1. Confirm the workbook path in the sidebar.
2. Pick the worksheet to use.
3. Add entries in **Quick Entry** and click **Save Entry**.
4. Search in **Search and View**.

## Notes

- If saving fails, close the workbook in Excel and try again.
- By default, backups are saved to a `PO_Backups` folder next to the workbook.
- If the workbook path does not exist, the app can create a new workbook with default columns.
- If a sheet is missing `Reason for Purchase`, POtrol adds it automatically on save.

## Build A Self-Contained EXE (Windows)

Use the included build script:

```powershell
cd "c:\Users\cholt\OneDrive - Champagne Metals\Documents\GitHub\POTracker"
.\build_exe.ps1
```

Or double-click:

- `build_exe.bat`

Output:

- `dist\POtrol.exe`

Notes:

- The script creates/uses `.venv-build` automatically.
- To use an existing Python environment instead, run:

```powershell
.\build_exe.ps1 -NoVenv
```
