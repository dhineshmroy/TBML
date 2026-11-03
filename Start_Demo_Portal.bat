@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" goto :noinstall
echo DEMO MODE - uses fake sample regions in the demo_data folder. Not for real reporting.
set TBML_DATA_DIR=demo_data
".venv\Scripts\python.exe" -m streamlit run app.py --browser.gatherUsageStats false
pause
exit /b 0
:noinstall
echo Please run Install_Portal.bat first.
pause
exit /b 1
