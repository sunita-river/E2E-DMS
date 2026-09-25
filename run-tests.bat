@echo off
REM Double-click to set up (installs anything missing) and then run all tests.
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1"
if errorlevel 1 (
  echo Setup failed - see messages above.
  pause
  exit /b 1
)
call npx playwright test %*
pause
