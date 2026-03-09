@echo off
setlocal

set "SCRIPT_DIR=%~dp0WindowLayout"
set "SCRIPT_PATH=%SCRIPT_DIR%\WindowLayout.ps1"

if not exist "%SCRIPT_PATH%" (
  echo Could not find "%SCRIPT_PATH%".
  echo Keep WindowLayout.cmd next to the WindowLayout folder.
  pause
  exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" %*
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
  echo.
  echo Script finished with exit code %EXIT_CODE%.
  pause
)

exit /b %EXIT_CODE%
