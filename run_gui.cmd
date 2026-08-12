@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    where py >nul 2>&1
    if errorlevel 1 (
        python -m venv .venv
    ) else (
        py -m venv .venv
    )
    if errorlevel 1 goto error
)

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto error

start "" ".venv\Scripts\pythonw.exe" "app.py"
exit /b 0

:error
echo.
echo Failed to start the application.
echo Please keep this window open and send a screenshot of the error above.
echo.
pause
exit /b 1
