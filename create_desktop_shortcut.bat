@echo off
setlocal
set "APP_DIR=%~dp0"
set "APP_EXE=%APP_DIR%dist\CometPDF.exe"

if not exist "%APP_EXE%" (
  echo CometPDF.exe was not found in the dist folder.
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\CometPDF.lnk'); $s.TargetPath='%APP_EXE%'; $s.WorkingDirectory='%APP_DIR%'; $s.Save()"

echo Desktop shortcut created: CometPDF
pause
