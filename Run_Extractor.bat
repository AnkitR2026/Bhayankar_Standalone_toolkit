@echo off
rem Bhayankar extractor (standalone): scans .\input and writes the digest, draft DAX and rebuild plan to .\output
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (
    python bhayankar_extractor.py %*
) else (
    py -3 bhayankar_extractor.py %*
)
echo.
pause
