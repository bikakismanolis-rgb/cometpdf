@echo off
setlocal EnableExtensions

set "INSTALL_DIR=%LOCALAPPDATA%\Programs\CometPDF"

reg delete "HKCU\Software\Classes\*\shell\CometPDF" /f >nul 2>nul
del "%USERPROFILE%\Desktop\CometPDF.lnk" >nul 2>nul
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\CometPDF.lnk" >nul 2>nul

if exist "%INSTALL_DIR%" rmdir /S /Q "%INSTALL_DIR%"

echo CometPDF removed.
pause
