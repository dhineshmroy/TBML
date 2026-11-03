@echo off
setlocal
cd /d "%~dp0"
echo ===== TBML Reporting Portal - installation =====
where python >nul 2>nul
if errorlevel 1 goto :nopython
if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment...
  python -m venv .venv
  if errorlevel 1 goto :fail
)
echo Installing packages ^(needs internet, or a "wheels" folder for offline install^)...
if exist wheels (
  ".venv\Scripts\python.exe" -m pip install --no-index --find-links wheels -r requirements.txt
) else (
  ".venv\Scripts\python.exe" -m pip install --upgrade pip
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
)
if errorlevel 1 goto :fail
echo Creating desktop shortcut...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath(Desktop)+\TBML Reporting Portal.lnk);$s.TargetPath=%~dp0Start_TBML_Portal.bat;$s.WorkingDirectory=%~dp0;$s.WindowStyle=7;$s.Save()"
echo.
echo Installation complete. Use the "TBML Reporting Portal" shortcut on the desktop.
pause
exit /b 0
:nopython
echo Python was not found. Install Python 3.11 or 3.12 ^(64-bit^) from python.org,
echo tick "Add python.exe to PATH", then run this file again.
pause
exit /b 1
:fail
echo Installation failed. Please send a screenshot of this window to the developer.
pause
exit /b 1
