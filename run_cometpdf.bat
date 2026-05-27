@echo off
setlocal
set "APP_DIR=%~dp0"
set "RUNTIME_PY=C:\Users\user\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if exist "%RUNTIME_PY%" (
  "%RUNTIME_PY%" "%APP_DIR%cometpdf_app.py"
) else (
  python "%APP_DIR%cometpdf_app.py"
)
