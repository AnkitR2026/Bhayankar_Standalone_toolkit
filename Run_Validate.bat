@echo off
rem Bhayankar validate: checks the PBIR files and the semantic model of any Power BI project offline (no Node.js, no internet).
rem Usage: Run_Validate.bat "path\to\Report.pbip"   (or drag a .pbip onto this file)
cd /d "%~dp0"
if "%~1"=="" (
    echo Drag a .pbip file onto Run_Validate.bat, or run: Run_Validate.bat "path\to\Report.pbip"
    pause
    exit /b 1
)
where python >nul 2>nul
if %errorlevel%==0 (
    python bhayankar_validate.py %* --format text
    python bhayankar_modelcheck.py %*
) else (
    py -3 bhayankar_validate.py %* --format text
    py -3 bhayankar_modelcheck.py %*
)
echo.
pause
