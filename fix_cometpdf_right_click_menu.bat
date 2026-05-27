@echo off
setlocal EnableExtensions

set "APP_EXE=%LOCALAPPDATA%\Programs\CometPDF\CometPDF.exe"

if not exist "%APP_EXE%" (
  echo CometPDF.exe was not found at:
  echo %APP_EXE%
  echo.
  echo Run CometPDFSetup-1.0.0.exe first, then run this file again.
  pause
  exit /b 1
)

reg add "HKCU\Software\Classes\*\shell\CometPDF" /ve /d "Convert to PDF" /f >nul
reg add "HKCU\Software\Classes\*\shell\CometPDF" /v "Icon" /d "%APP_EXE%" /f >nul
reg add "HKCU\Software\Classes\*\shell\CometPDF\command" /ve /d "\"%APP_EXE%\" --convert \"%%1\" --open" /f >nul

echo Right-click conversion was updated.
echo PDFs will be saved next to the original file and opened automatically.
pause
