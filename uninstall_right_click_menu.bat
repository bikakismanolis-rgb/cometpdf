@echo off
reg delete "HKCU\Software\Classes\*\shell\CometPDF" /f
echo Removed CometPDF right-click menu.
pause
