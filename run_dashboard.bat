@echo off
rem Generate / refresh the IBGT turnover dashboard (outputs land next to this file)
cd /d "%~dp0"
if exist make_hr_dashboard_standalone.py (
  python make_hr_dashboard_standalone.py
) else (
  python make_hr_dashboard.py
)
echo.
echo If you see "No module named openpyxl", run:  pip install openpyxl
pause
