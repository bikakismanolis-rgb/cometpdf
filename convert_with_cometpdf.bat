@echo off
setlocal
set "APP_DIR=%~dp0"
set "APP_EXE=%APP_DIR%dist\CometPDF.exe"
set "RUNTIME_PY=C:\Users\user\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if "%~1"=="" (
  echo Drop a file on this script or use the Windows right-click menu.
  pause
  exit /b 1
)

if exist "%APP_EXE%" (
  "%APP_EXE%" --convert "%~1" --open
) else if exist "%RUNTIME_PY%" (
  "%RUNTIME_PY%" "%APP_DIR%cometpdf_app.py" --convert "%~1" --open
) else (
  python "%APP_DIR%cometpdf_app.py" --convert "%~1" --open
)

if errorlevel 1 (
  pause
)
