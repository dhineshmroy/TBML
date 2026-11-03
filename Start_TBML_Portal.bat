@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" goto :noinstall
echo Starting TBML Reporting Portal... keep this window open while you work.
".venv\Scripts\python.exe" -m streamlit run app.py --browser.gatherUsageStats false
pause
exit /b 0
:noinstall
echo Please run Install_Portal.bat first.
pause
exit /b 1
