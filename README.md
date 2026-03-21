# 9-Grid Talent Calibration Generator

Python script to generate a clean 9-grid talent calibration chart from Excel.

The script reads an Excel workbook, plots a 3x3 performance/trajectory grid, exports a high-resolution PNG, and creates CSV tables for discussion. It also generates one chart per owner/manager when enough data is available.

## What It Produces

- `output/9grid_overview.png`
- `output/9grid_overview_legend.csv`
- `output/9grid_owner_<owner>.png`
- `output/9grid_owner_<owner>_legend.csv`

## Requirements

- Python 3.10+
- `pandas`
- `matplotlib`
- `openpyxl`

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Expected Excel File

By default, the script looks for either of these files:

- `9Grid exercice.xlsx`
- `9Grid_exercice.xlsx`

in the same folder as the script.

It reads the sheet named:

- `Data`

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

- `Brecht` = blue
- `Gary` = green
- `Stephanie` = purple
- `Bart` = teal

Other owners receive fallback colors automatically.

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

## How To Run

From the project folder:

```powershell
python .\generate_9grid.py
```

If `python` does not work on your machine:

```powershell
py .\generate_9grid.py
```

## Notes

- Close the Excel file before running the script, otherwise Windows may block access.
- Generated files are written to the `output/` folder.
- The input workbook and generated output are ignored by Git in the default setup.
