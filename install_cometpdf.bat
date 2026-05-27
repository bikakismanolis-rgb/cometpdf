@echo off
setlocal EnableExtensions

set "APP_NAME=CometPDF"
set "INSTALL_DIR=%LOCALAPPDATA%\Programs\CometPDF"
set "SOURCE_DIR=%~dp0"
set "EXE=%INSTALL_DIR%\CometPDF.exe"

echo Installing %APP_NAME%...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

copy /Y "%SOURCE_DIR%CometPDF.exe" "%INSTALL_DIR%\" >nul
copy /Y "%SOURCE_DIR%README.md" "%INSTALL_DIR%\" >nul
copy /Y "%SOURCE_DIR%THIRD_PARTY_NOTICES.md" "%INSTALL_DIR%\" >nul
copy /Y "%SOURCE_DIR%cometpdf.ico" "%INSTALL_DIR%\" >nul

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $desktop=[Environment]::GetFolderPath('Desktop'); $start=[Environment]::GetFolderPath('StartMenu') + '\Programs'; $targets=@($desktop + '\CometPDF.lnk', $start + '\CometPDF.lnk'); foreach($target in $targets){ $s=$ws.CreateShortcut($target); $s.TargetPath='%EXE%'; $s.WorkingDirectory='%INSTALL_DIR%'; $s.IconLocation='%INSTALL_DIR%\cometpdf.ico'; $s.Save() }"

choice /C YN /N /M "Add 'Convert to PDF' to Windows right-click menu? [Y/N] "
if errorlevel 2 goto skip_context

reg add "HKCU\Software\Classes\*\shell\CometPDF" /ve /d "Convert to PDF" /f >nul
reg add "HKCU\Software\Classes\*\shell\CometPDF" /v "Icon" /d "%INSTALL_DIR%\cometpdf.ico" /f >nul
reg add "HKCU\Software\Classes\*\shell\CometPDF\command" /ve /d "\"%EXE%\" --convert \"%%1\" --open" /f >nul

:skip_context
echo.
echo CometPDF installed successfully.
echo Location: %INSTALL_DIR%
pause
