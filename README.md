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

## 📝 Input

Place the Excel export file from whatever tracking system you use in the repository root and name it `Space Systems Division.xlsx` (or adjust the script accordingly).

## 🚀 Running the report

Use the provided helper script (PowerShell):

```powershell
./run_report.ps1
```

or run directly:

```powershell
python weekly_report.py
```

The script will create a `reports/` directory and write a timestamped workbook containing:

* **Metrics** – per-person counts for late tasks, urgent items, etc.
* **Late Details** – list of overdue items
* **Done Last Week** – tasks completed in the previous 7 days
* **Planned Next Week** – tasks due in the coming week

It also prints text summaries to the console.

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

(Include whichever license applies or remove this section.)

---

Feel free to modify the script to fit your team's workflow or to convert the input source. Questions or improvements welcome!