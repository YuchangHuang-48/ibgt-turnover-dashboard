@echo off
rem 双击运行：在脚本所在文件夹生成/更新 IBGT 人员流动看板
chcp 65001 >nul
cd /d "%~dp0"
python make_hr_dashboard.py
echo.
echo 如提示 No module named openpyxl，请先执行:
echo     pip install openpyxl
pause
