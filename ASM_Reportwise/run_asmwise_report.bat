@echo off
rem ASM-wise report: downloads the 6 MIS exports into ASM_Reportwise\Downloads, builds
rem ASM_Reportwise\Output\ASMwise_AllSources_Summary_<yesterday>.xlsx and, if that worked, mails
rem each ASM in ASM_Reportwise\Mapping\ASM_Name_Mapping.xlsx. For a local run; the schedule is
rem .github\workflows\asmwise-report.yml (daily at 08:00 IST).
rem Output goes to ASM_Reportwise\Output\logs\asmwise_report.log.
cd /d "%~dp0.."
set LOG=ASM_Reportwise\Output\logs\asmwise_report.log
if not exist "ASM_Reportwise\Output\logs" mkdir "ASM_Reportwise\Output\logs"
set PYTHONIOENCODING=utf-8
echo ===== %DATE% %TIME% ===== >> "%LOG%"
"C:\Python314\python.exe" ASM_Reportwise\asmwise_report.py >> "%LOG%" 2>&1
set RC=%ERRORLEVEL%
echo report exit code %RC% >> "%LOG%"
if not %RC%==0 exit /b %RC%
"C:\Python314\python.exe" ASM_Reportwise\asm_email.py --send >> "%LOG%" 2>&1
set RC=%ERRORLEVEL%
echo email exit code %RC% >> "%LOG%"
exit /b %RC%
