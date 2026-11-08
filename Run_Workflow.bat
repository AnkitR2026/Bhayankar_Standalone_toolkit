@echo off
rem Bhayankar workflow: for every workbook in input\ that is new or changed, run extract, gate, build, validate and test.
rem Result: output\<workbook>\WORKFLOW_STATUS.md  (read this first)
rem Options: --no-build  --no-test  --force   or a workbook name:  Run_Workflow.bat "My Report"
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (
    python bhayankar_run.py run %*
) else (
    py -3 bhayankar_run.py run %*
)
pause
