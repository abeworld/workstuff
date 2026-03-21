# 9-Grid Talent Calibration Generator

Python script to generate a clean 9-grid talent calibration chart from Excel.

The script reads one or more Excel workbooks, plots a 3x3 performance/trajectory grid, exports a high-resolution PNG, creates CSV tables for discussion, and generates one Power BI-ready workbook per manager input. It also generates one chart per owner/manager when enough data is available.

## What It Produces

- `output/9grid_overview.png`
- `output/9grid_overview_legend.csv`
- `output/9grid_owner_<owner>.png`
- `output/9grid_owner_<owner>_legend.csv`
- `output/9GRID_<manager>.xlsb`

## Requirements

- Python 3.10+
- `pandas`
- `matplotlib`
- `openpyxl`

Install dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Expected Excel Files

Preferred input format:

- `9grid_Gary.xlsx`
- `9grid_Brecht.xlsx`
- `9grid_Stephanie.xlsx`

Pattern:

- `9grid_<manager>.xlsx`

You can place multiple files in the same folder and the script will process all of them in one run.

How manager naming works:

- the manager name is taken from the filename
- `9grid_Gary.xlsx` becomes manager `Gary`
- `9grid_Bart_Vandenberghe.xlsx` becomes manager `Bart Vandenberghe`
- all rows loaded from that workbook are assigned to that manager for color-coding and owner-specific exports
- that workbook also produces a Power BI export named `9GRID_<manager>.xlsb`

Fallback for older usage:

- `9Grid exercice.xlsx`
- `9Grid_exercice.xlsx`

All input workbooks must be in the same folder as the script.

Each workbook is expected to read the sheet named:

- `Data`

The Power BI template workbook is expected at one of:

- `.\9GRID.xlsb`
- `C:\Users\<you>\Downloads\9GRID.xlsb`

## Supported Input Formats

The script currently supports three Excel layouts.

### 1. Compact format

Required columns:

- `Lead Name`
- `Team Member`
- `9Grid number`

Optional columns:

- `Action Bucket`
- `Risk of Churn (1->3 - Low medium high)`

How compact format is interpreted:

- `Lead Name` becomes the owner/manager
- `9Grid number` is converted into the 3x3 grid position
- `Risk of Churn (1->3 - Low medium high)` is mapped to churn styling

### 2. Detailed format

Required columns:

- `Team Member`
- `Action Bucket`
- `Owner`
- `Churn Risk`
- `Main Strength`
- `Main Concern`
- `Rationale`
- `Performance Score`
- `Trajectory Score`

### 3. Hybrid calibration export

This matches the newer workbook structure that includes HR data plus calibration fields.

Required columns:

- `Lead Name`
- `Name`
- `Action Bucket`
- `Owner`
- `Main Strength`
- `Main Concern`
- `Rationale`
- `Churn Risk`
- `Performance`
- `Potential`

How hybrid format is interpreted:

- `Name` becomes the plotted team member name
- `Performance` is mapped to `Performance Score`
- `Potential` is mapped to `Trajectory Score`
- `Owner` is used when filled; otherwise the script falls back to `Lead Name`
- rows without valid `Performance` and `Potential` scores from `1` to `3` are skipped
- a template/example row named `John Doe` is always ignored

## Visual Encoding

### Position

- X-axis = current performance
  - `1` = Below bar
  - `2` = At bar
  - `3` = Above bar
- Y-axis = growth trajectory / ownership
  - `1` = Stable
  - `2` = Growable
  - `3` = High trajectory

### Manager colors

- each manager gets a distinct color within a run
- colors may change between runs
- manager colors are derived from the filenames when using `9grid_<manager>.xlsx`

### Churn risk styling

- `1` or `Low` = lighter treatment
- `2` or `Medium` = stronger orange outline and larger marker
- `3` or `High` = stronger red outline and larger marker

## Overview Table

The overview export groups by team member and shows:

- employee number
- team member name
- one score column per manager

## Owner-Specific Charts

The script creates owner-specific charts when an owner has at least 2 valid plotted employees.

## Power BI Export

For each input workbook, the script creates one `.xlsb` output based on the `9GRID.xlsb` template.

The generated workbook keeps the template structure and fills the `9grid` sheet with exactly these columns:

- `Department`
- `Employee ID`
- `Name`
- `9Grid_Date`
- `Flight Risk`
- `Performance`
- `Potential`
- `Grid Box`
- `Feedback`

Value mapping:

- `Name` comes from the employee name in the input workbook
- `Flight Risk`, `Performance`, and `Potential` are exported as `Low`, `Moderate`, or `High`
- `Grid Box` is exported as combinations such as `Low-High` or `Moderate-Moderate`
- `Feedback` uses the input `Rationale` when no dedicated feedback column exists
- a template/example row named `John Doe` is always excluded

## How To Run

From the project folder:

```powershell
.\.venv\Scripts\Activate.ps1
python .\generate_9grid.py
```

If your terminal does not recognize `python` yet, run it directly from the virtual environment:

```powershell
.\.venv\Scripts\python.exe .\generate_9grid.py
```

## Notes

- Close the Excel file before running the script, otherwise Windows may block access.
- When multiple `9grid_<manager>.xlsx` files are present, they are merged into one overview and split back out into one owner chart per manager.
- A row for `John Doe` is treated as template guidance and is excluded from all outputs.
- One Power BI-ready `.xlsb` file is generated per input workbook in the `output/` folder.
- Generated files are written to the `output/` folder.
- `.venv/`, the input workbook, and generated output are ignored by Git in the default setup.
