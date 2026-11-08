@echo off
rem Bhayankar build (standalone): builds output\<workbook>\pbip\<workbook>.pbip from the workbooks in .\input
rem The extractor runs first by itself when there is no extract yet or the input changed since the last one
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (
    python bhayankar_build.py %*
) else (
    py -3 bhayankar_build.py %*
)
echo.
pause
