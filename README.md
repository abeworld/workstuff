# 9-Grid Talent Calibration Generator

This project processes one or more manager Excel files, builds 9-grid calibration visuals, and creates Power BI-ready `.xlsb` exports.

## What It Does

For each run, the script:

- reads all input workbooks matching `9grid_<manager>.xlsx`
- combines them into one overview dataset
- generates an overview 9-grid PNG
- generates one owner-specific 9-grid PNG when that owner has at least 2 valid employees
- generates one Power BI-ready `9GRID_<manager>.xlsb` file per input workbook

## Expected Input Files

Place the input files in the same folder as [generate_9grid.py](/C:/Users/goldm/Downloads/Codex_work/workstuff/generate_9grid.py).

Preferred pattern:

- `9grid_Gary.xlsx`
- `9grid_Brecht.xlsx`
- `9grid_Stephanie.xlsx`

Naming rule:

- `9grid_<manager>.xlsx`

Examples:

- `9grid_Gary.xlsx` maps to manager `Gary`
- `9grid_Bart_Vandenberghe.xlsx` maps to manager `Bart Vandenberghe`

All rows loaded from that workbook are assigned to that manager for:

- manager color mapping in the PNGs
- owner-specific exports
- Power BI workbook naming

Fallback for older single-file usage is still supported:

- `9Grid exercice.xlsx`
- `9Grid_exercice.xlsx`

## Excel Sheet Requirements

Each input workbook must contain a sheet named:

- `Data`

The script supports these input layouts.

### Compact Format

Required columns:

- `Lead Name`
- `Team Member`
- `9Grid number`

Optional columns:

- `Action Bucket`
- `Risk of Churn (1->3 - Low medium high)`

### Detailed Format

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

### Hybrid Calibration Format

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

Interpretation rules:

- `Name` becomes the plotted employee name
- `Performance` maps to `Performance Score`
- `Potential` maps to `Trajectory Score`
- blank `Owner` falls back to `Lead Name`
- duplicate imported columns such as `Rationale.1` are merged and the filled value is kept
- rows without valid performance and potential scores from `1` to `3` are skipped
- a row named `John Doe` is always ignored

## Output Files

Generated files are written to [C:\Users\goldm\Downloads\Codex_work\workstuff\output](C:\Users\goldm\Downloads\Codex_work\workstuff\output).

Typical outputs:

- `output/9grid_overview.png`
- `output/9grid_owner_<owner>.png`
- `output/9GRID_<manager>.xlsb`

## PNG Output Behavior

### Overview PNG

The overview chart combines all valid employees from all input files.

The legend panel includes:

- manager colors
- churn risk outline guide
- one row per employee
- one score column per manager

Current layout behavior:

- the overview table uses `Nr` as the first column header
- `Manager colors` and `Churn risk outline` are shown next to each other to leave more room for the table

### Owner-Specific PNG

An owner-specific chart is created only when that owner has at least 2 valid employees.

Current layout behavior:

- the owner legend table uses `Nr` as the first column header
- the `Owner` column is intentionally removed because the chart already belongs to that owner

### Visual Encoding

Position:

- X-axis = current performance
- Y-axis = growth trajectory / ownership

Score interpretation:

- `1` = low / below
- `2` = moderate / at bar
- `3` = high / above

Churn risk styling:

- `Low` = lighter outline
- `Medium` = stronger orange outline
- `High` = stronger red outline

Manager colors:

- each manager gets a distinct color within a run
- colors may change between runs

## Power BI Export

For each input workbook, the script creates one Power BI-ready workbook based on a template named `9GRID.xlsb`.

Expected template locations:

- `.\9GRID.xlsb`
- `C:\Users\<you>\Downloads\9GRID.xlsb`

Generated workbook naming:

- `9GRID_Gary.xlsb`
- `9GRID_Brecht.xlsb`

The script keeps the template workbook structure and fills the `9grid` sheet with exactly these columns:

- `Department`
- `Employee ID`
- `Name`
- `9Grid_Date`
- `Flight Risk`
- `Performance`
- `Potential`
- `Grid Box`
- `Feedback`

Mapping rules:

- `Name` comes from the employee name in the input workbook
- `Flight Risk` is exported as `Low`, `Moderate`, or `High`
- `Performance` is exported as `Low`, `Moderate`, or `High`
- `Potential` is exported as `Low`, `Moderate`, or `High`
- `Grid Box` is written as the Excel formula from the template logic, not as hardcoded text
- `Feedback` uses the input `Feedback` field when present and otherwise falls back to `Rationale`
- if rationale text exists in a duplicate imported column such as column `W` / `Rationale.1`, that value is used
- `John Doe` is excluded from the export

## Requirements

- Python 3.10+
- `pandas`
- `matplotlib`
- `openpyxl`
- `pywin32`

Install everything with:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## How To Run

From PowerShell:

```powershell
cd C:\Users\goldm\Downloads\Codex_work\workstuff
.\.venv\Scripts\Activate.ps1
python .\generate_9grid.py
```

If `python` is not recognized in your terminal:

```powershell
cd C:\Users\goldm\Downloads\Codex_work\workstuff
.\.venv\Scripts\python.exe .\generate_9grid.py
```

## Full Workflow

1. Drop one or more `9grid_<manager>.xlsx` files into the project folder.
2. Make sure the `9GRID.xlsb` Power BI template is available.
3. Run `generate_9grid.py`.
4. Pick up the generated PNG and `.xlsb` files from the `output` folder.

## Notes

- Close Excel files before running the script, otherwise Windows may block access.
- When multiple manager files are present, they are merged into one overview.
- Owner-specific PNGs are generated from the combined dataset after manager assignment.
- `.venv/`, input workbooks, and generated output are ignored by Git by default.
