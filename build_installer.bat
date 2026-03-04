@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"

if not exist .venv\Scripts\activate.bat (
  echo [ERROR] venv not found. Run: python -m venv .venv
  pause
  exit /b 1
)

call .venv\Scripts\activate

rmdir /s /q build dist 2>nul
del /q "VARDA Control Center.spec" 2>nul

set MISSING=0
for %%F in (img_logo.png img_icon1.png img_icon2.png img_icon3.png) do (
  if not exist "%%F" (
    echo [ERROR] Missing file: %%F
    set MISSING=1
  )
)

if "!MISSING!"=="1" (
  echo.
  echo Put the PNGs next to app.py, or rename them to the expected names.
  pause
  exit /b 1
)

pyinstaller --noconfirm --clean --windowed --name "VARDA Control Center" ^
  --icon varda.ico ^
  --add-data "varda.ico;." ^
  --add-data "img_logo.png;." ^
  --add-data "img_icon1.png;." ^
  --add-data "img_icon2.png;." ^
  --add-data "img_icon3.png;." ^
  app.py

if errorlevel 1 (
  echo.
  echo [ERROR] PyInstaller build failed.
  pause
  exit /b 1
)

echo.
echo ✅ Build done:
echo dist\VARDA Control Center\VARDA Control Center.exe
pause