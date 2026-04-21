@echo off
rem pdf/UA launcher - Windows.
rem Uses LibreOffice's bundled Python (it has `uno` built in).

setlocal enabledelayedexpansion
cd /d "%~dp0"

rem Step 1. Locate LibreOffice python.exe. System Python on Windows
rem cannot `import uno` without extra setup; the bundled one can.
set "LOPY="
set "LOBIN="
for %%D in (
  "C:\Program Files\LibreOffice\program"
  "C:\Program Files (x86)\LibreOffice\program"
  "%LOCALAPPDATA%\Programs\LibreOffice\program"
) do (
  if exist "%%~D\python.exe" (
    if "!LOPY!"=="" (
      set "LOPY=%%~D\python.exe"
      set "LOBIN=%%~D"
    )
  )
)

if "!LOPY!"=="" (
  echo [pdf/UA] ERROR: LibreOffice not found in standard locations.
  echo           Install LibreOffice from https://www.libreoffice.org/
  echo           or set LOPY manually in this script.
  pause
  exit /b 1
)

echo [pdf/UA] using LibreOffice Python: !LOPY!
set "PATH=!LOBIN!;%PATH%"

rem Step 2. Make sure pip is available in the bundled Python.
"!LOPY!" -m pip --version >nul 2>nul
if errorlevel 1 (
  echo [pdf/UA] bootstrapping pip in LibreOffice Python...
  "!LOPY!" -m ensurepip --default-pip
  if errorlevel 1 (
    echo [pdf/UA] ERROR: could not bootstrap pip. Run as administrator?
    pause
    exit /b 1
  )
)

rem Step 3. Install our pip deps if missing.
"!LOPY!" -c "import flask, PIL, pytesseract" >nul 2>nul
if errorlevel 1 (
  echo [pdf/UA] installing Python dependencies...
  "!LOPY!" -m pip install --user --quiet -r requirements.txt
  if errorlevel 1 (
    echo [pdf/UA] ERROR: pip install failed.
    pause
    exit /b 1
  )
)

rem Step 4. Sanity checks.
"!LOPY!" -c "import uno" >nul 2>nul
if errorlevel 1 (
  echo [pdf/UA] ERROR: LibreOffice Python cannot import uno. Broken install?
  pause
  exit /b 1
)

where tesseract >nul 2>nul
if errorlevel 1 (
  echo [pdf/UA] WARNING: tesseract.exe not found in PATH. OCR will be disabled.
  echo           Install from https://github.com/UB-Mannheim/tesseract/wiki
  echo           and add its folder to PATH.
)

if "%PDFUA_PORT%"=="" set PDFUA_PORT=8000
if "%PDFUA_HOST%"=="" set PDFUA_HOST=127.0.0.1

echo.
echo [pdf/UA] starting on http://%PDFUA_HOST%:%PDFUA_PORT%/
echo [pdf/UA] keep this window open while using the app. Ctrl+C to stop.
echo.

"!LOPY!" -m pdfua.cli serve --host %PDFUA_HOST% --port %PDFUA_PORT% --open

rem If the server exited (normally via Ctrl+C or with an error), pause so
rem the user can read any traceback before the cmd window closes.
echo.
echo [pdf/UA] server exited.
pause
