@echo off
rem Power BI report test (standalone): checks each Power BI project (PBIP) in .\test_input\<Name>\ against its Tableau workbook
rem and writes .\test_output\<Name>\test_report.xlsx (see README section 3)
cd /d "%~dp0"
if not exist "test_input" (
    mkdir "test_input"
    echo Created the test_input folder. Put one folder per report in it:
    echo   test_input\^<Name^>\^<Name^>.twb  plus the Power BI project ^<Name^>.pbip with its .SemanticModel and .Report folders
    echo Then run Run_PowerBI_Report_Test.bat again.
    echo.
    pause
    exit /b 1
)
where python >nul 2>nul
if %errorlevel%==0 (
    python bhayankar_test.py %*
) else (
    py -3 bhayankar_test.py %*
)
set RC=%errorlevel%
echo.
if %RC%==0 (echo RESULT: done - open test_output\^<Name^>\test_report.xlsx) else (echo RESULT: the compare stopped with an error - see output above)
pause
exit /b %RC%
