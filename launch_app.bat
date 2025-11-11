@echo off
setlocal ENABLEDELAYEDEXPANSION
set "APP_DIR=%~dp0"
cd /d "%APP_DIR%"
set "PORT=8501"
set "URL=http://localhost:%PORT%/"
if not exist "venv" (
  echo [setup] Creating Python venv...
  py -m venv "venv" || goto :fail
  echo [setup] Upgrading pip...
  call "venv\Scripts\python.exe" -m pip install --upgrade pip || goto :fail
  echo [setup] Installing requirements...
  call "venv\Scripts\pip.exe" install -r "requirements.txt" || goto :fail
)
echo [run] Starting server on %URL%
start "" "%APP_DIR%venv\Scripts\pythonw.exe" -m streamlit run "app.py" ^
  --server.port %PORT% ^
  --server.headless true ^
  --browser.gatherUsageStats false
timeout /t 3 >nul
start "" "%URL%"
exit /b 0
:fail
echo.
echo *** Something went wrong. ***
pause
exit /b 1
