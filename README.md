# Weekly Task Report

This repository contains a small Python script that generates a weekly status report for the Space Systems Division team by processing an Excel export.

## 📁 Project Structure

```
weekly_report.py          # main script that reads Excel input and writes an output workbook
run_report.ps1           # helper PowerShell script to invoke the report
dist/                    # (if used) PyInstaller output
reports/                 # generated Excel reports (ignored by Git)
Space Systems Division.xlsx  # input spreadsheet (not tracked by Git)
weekly_report.spec       # PyInstaller spec file for building an executable
.venv/                   # Python virtual environment
.gitignore               # rules to keep build/output files out of source control
```

## 🔧 Requirements

- Python 3.11 or later
- `pandas` and `openpyxl` (install via `pip install -r requirements.txt`)

> A virtual environment is highly recommended. You can create one with:
>
> ```powershell
> python -m venv .venv
> .\.venv\Scripts\Activate.ps1
> pip install -r requirements.txt
> ```

## � Input data (what the script expects)

The script reads an Excel workbook (`.xlsx`) and expects the following columns to exist (case-sensitive):

- `Task Name` – the title or short description of the task.
- `Bucket Name` – used to classify tasks (e.g., `Tasks`, `Ready For Review`, `Completed`). Rows with `Bucket Name` containing `Blocked`, `Archive`, `Archived`, or `Discontinued` are ignored.
- `Assigned To` – a semicolon-separated list of assignees (e.g., `Alice; Bob`). The script expands multi‑assignee cells and currently filters to a small team by name.
- `Due date` – a date/datetime used to identify overdue and upcoming work.
- `Completed Date` – a date/datetime used to identify what was finished in the last week.
- `Progress` – used to ignore items already marked as `Completed`.
- `Priority` – used to identify `Urgent` items.


## 📁 Where to put the data files

Place one or more Excel exports (`.xlsx` files) in the repository root (next to `weekly_report.py`).

**All** `.xlsx` files in the directory will be processed automatically. Each will generate a separate report in the `reports/` folder, named after the input file.


## 🚀 Running the report

Use the provided helper script (PowerShell):

```powershell
./run_report.ps1
```

or run directly:

```powershell
python weekly_report.py
```

The script will process every `.xlsx` file in the directory, creating a separate report for each in the `reports/` folder. Each output file will be named after the input file (e.g., `MyInput.xlsx` → `reports/MyInput_weekly_report_YYYY-MM-DD.xlsx`).

Each report contains:

* **Metrics** – per-person counts for late tasks, urgent items, etc.
* **Late Details** – list of overdue items
* **Done Last Week** – tasks completed in the previous 7 days
* **Planned Next Week** – tasks due in the coming week

Text summaries are printed to the console for each input file.

## 🛠️ Building an executable

A `weekly_report.spec` file is provided for PyInstaller. To generate a standalone `.exe`:

```powershell
pyinstaller --onefile weekly_report.spec
```

Output will appear under `dist/` and `build/`.

## ✅ Git Best Practices

- Only commit source code; generated files are ignored by `.gitignore`.
- Build artifacts and report outputs live in `build/`, `dist/`, or `reports/` which are not tracked.
- Use clear, small commits and descriptive messages.

## 📎 License

Private
