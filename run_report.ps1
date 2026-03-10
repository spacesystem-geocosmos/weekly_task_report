# Run the weekly report and place output in ./reports/
# Usage: ./run_report.ps1

$ErrorActionPreference = 'Stop'
python .\weekly_report.py
