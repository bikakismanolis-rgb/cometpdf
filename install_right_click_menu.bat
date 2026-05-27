@echo off
setlocal
set "APP_DIR=%~dp0"
set "COMMAND=%APP_DIR%convert_with_cometpdf.bat"

reg add "HKCU\Software\Classes\*\shell\CometPDF" /ve /d "Convert to PDF" /f
reg add "HKCU\Software\Classes\*\shell\CometPDF" /v "Icon" /d "%APP_DIR%cometpdf.ico" /f
reg add "HKCU\Software\Classes\*\shell\CometPDF\command" /ve /d "\"%COMMAND%\" \"%%1\"" /f

echo Added right-click menu: Convert to PDF
pause
