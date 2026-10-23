@echo off
rem ASM-wise report: downloads the 6 MIS exports into Output\Downloads\ASM, builds
rem Output\ASM\ASMwise_AllSources_Summary_<yesterday>.xlsx and, if that worked, mails
rem each ASM in asm-report\Mapping\ASM_Name_Mapping.xlsx. For a local run; the schedule is
rem .github\workflows\asmwise-report.yml (daily at 08:00 IST).
rem Output goes to Output\ASM\logs\asmwise_report.log.
cd /d "%~dp0.."
set LOG=Output\ASM\logs\asmwise_report.log
if not exist "Output\ASM\logs" mkdir "Output\ASM\logs"
set PYTHONIOENCODING=utf-8
echo ===== %DATE% %TIME% ===== >> "%LOG%"
"C:\Python314\python.exe" asm-report\asmwise_report.py >> "%LOG%" 2>&1
set RC=%ERRORLEVEL%
echo report exit code %RC% >> "%LOG%"
if not %RC%==0 exit /b %RC%
"C:\Python314\python.exe" asm-report\asm_email.py --send >> "%LOG%" 2>&1
set RC=%ERRORLEVEL%
echo email exit code %RC% >> "%LOG%"
exit /b %RC%
