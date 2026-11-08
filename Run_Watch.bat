@echo off
rem Bhayankar watcher: runs the whole workflow (extract, gate, build, test) when a workbook lands in input\
rem Leave this window open while you work. Ctrl+C stops it. Result: output\<workbook>\WORKFLOW_STATUS.md
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (
    python bhayankar_watch.py %*
) else (
    py -3 bhayankar_watch.py %*
)
pause
