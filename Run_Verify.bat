@echo off
rem Bhayankar verify: checks a built project inside Power BI Desktop's own engine (model errors, rows, every measure, numbers vs Tableau).
rem Needs Power BI Desktop installed. Usage: Run_Verify.bat "My Report"   (add --open to open the project first, --refresh to load the data)
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (
    python bhayankar_verify.py %*
) else (
    py -3 bhayankar_verify.py %*
)
echo.
pause
