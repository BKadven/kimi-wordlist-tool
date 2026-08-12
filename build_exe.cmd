@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo ================================================
echo Building Kimi Wordlist Tool for Windows
echo ================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    where py >nul 2>&1
    if errorlevel 1 (
        python -m venv .venv
    ) else (
        py -m venv .venv
    )
    if errorlevel 1 goto error
)

echo [1/3] Installing build dependencies...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto error
".venv\Scripts\python.exe" -m pip install -r requirements-build.txt
if errorlevel 1 goto error

echo [2/3] Building EXE...
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean "KimiWordlist.spec"
if errorlevel 1 goto error

echo [3/3] Finished.
echo.
echo EXE location:
echo %~dp0dist\Kimi雅思单词本整理器_太阳版.exe
echo.
start "" "%~dp0dist"
pause
exit /b 0

:error
echo.
echo Build failed.
echo Please keep this window open and send a screenshot of the error above.
echo.
pause
exit /b 1
