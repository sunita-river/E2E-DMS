@echo off
REM Double-click to run the PDI blank-details fix-up (tests/PDI.spec.ts) in a visible browser.
REM Dealer comes from PIDUsers in resources/credentials.json (or set PDI_DEALER beforehand).
REM   run-pdi.bat        asks: dry run or save
REM   run-pdi.bat dry    fill the form only, no Save
REM   run-pdi.bat save   fill the form and click Save
cd /d "%~dp0"

set "MODE=%~1"
if "%MODE%"=="" (
  echo   D = Dry run  - fill the form, do NOT click Save
  echo   S = Save     - fill the form and click Save
  choice /c DS /n /m "Dry run or Save? [D/S] "
  if errorlevel 2 (set "MODE=save") else (set "MODE=dry")
)

if /i "%MODE%"=="dry" (
  set "PDI_DRY_RUN=1"
  echo Running PDI as a DRY RUN - nothing will be saved.
) else if /i "%MODE%"=="save" (
  set "PDI_DRY_RUN="
  echo Running PDI and SAVING changes in the DMS.
) else (
  echo Unknown option "%MODE%" - use dry or save.
  pause
  exit /b 1
)

set "PDI_RUN=1"
call npx playwright test tests/PDI.spec.ts --headed
pause
