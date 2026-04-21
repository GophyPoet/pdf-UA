@echo off
rem pdf/UA launcher - Windows.
rem Double-click in Explorer or run.bat from cmd/PowerShell.
rem Starts local Flask server and opens the UI in default browser.

setlocal
cd /d "%~dp0"

set PY=python
where %PY% >nul 2>nul
if errorlevel 1 (
  set PY=py -3
)

%PY% -c "import flask, PIL, pytesseract" >nul 2>nul
if errorlevel 1 (
  echo [pdf/UA] installing Python dependencies...
  %PY% -m pip install --quiet -r requirements.txt
)

where soffice >nul 2>nul
if errorlevel 1 (
  echo [pdf/UA] WARNING: soffice not found in PATH. Install LibreOffice and add it to PATH.
)
where tesseract >nul 2>nul
if errorlevel 1 (
  echo [pdf/UA] WARNING: tesseract not found in PATH. OCR will be disabled.
)

if "%PDFUA_PORT%"=="" set PDFUA_PORT=8000
if "%PDFUA_HOST%"=="" set PDFUA_HOST=127.0.0.1

echo [pdf/UA] starting on http://%PDFUA_HOST%:%PDFUA_PORT%/
%PY% -m pdfua.cli serve --host %PDFUA_HOST% --port %PDFUA_PORT% --open
