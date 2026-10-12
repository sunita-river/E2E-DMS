@echo off
rem Daily ASM-wise Online report: downloads the 6 MIS exports and builds
rem Output\ASMwise_Online_Summary_<yesterday>.xlsx. Run by Task Scheduler ("ASMwise Online Report").
rem Output goes to Output\logs\asmwise_report.log.
cd /d "%~dp0.."
if not exist "Output\logs" mkdir "Output\logs"
set PYTHONIOENCODING=utf-8
echo ===== %DATE% %TIME% ===== >> "Output\logs\asmwise_report.log"
"C:\Python314\python.exe" scripts\asmwise_report.py >> "Output\logs\asmwise_report.log" 2>&1
echo exit code %ERRORLEVEL% >> "Output\logs\asmwise_report.log"
exit /b %ERRORLEVEL%
