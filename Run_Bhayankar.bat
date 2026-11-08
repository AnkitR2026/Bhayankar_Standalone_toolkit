@echo off
rem Bhayankar menu: one launcher for every tool. Each choice runs the matching Run_*.bat (they also work on their own).
cd /d "%~dp0"
:menu
cls
echo.
echo   Bhayankar Standalone Toolkit
echo   ============================
echo   Put the Tableau workbook(s) in the input folder first (see README section 2).
echo.
echo   1  Extract    read the workbooks in input\ (digest, draft DAX)
echo   2  Build      build the Power BI project(s) and validate them
echo   3  Workflow   extract, check, build and test in one go
echo   4  Verify     check a built project inside Power BI Desktop
echo   5  Test       compare a Power BI project with its Tableau workbook (test_input\)
echo   6  Validate   check any .pbip's report files
echo   7  Self-check test the toolkit on this PC
echo   8  Open the output folder
echo   9  Open the README
echo   0  Exit
echo.
choice /c 1234567890 /n /m "  Choose 0-9: "
set C=%errorlevel%
if %C%==10 exit /b 0
if %C%==1 call Run_Extractor.bat
if %C%==2 call Run_Build.bat
if %C%==3 call Run_Workflow.bat
if %C%==4 goto verify
if %C%==5 call Run_PowerBI_Report_Test.bat
if %C%==6 goto validate
if %C%==7 call Run_SelfCheck.bat
if %C%==8 (if not exist output mkdir output) & start "" "%~dp0output"
if %C%==9 start "" "%~dp0README.md"
goto menu

:verify
echo.
echo   Workbook name = the file name without .twb/.twbx (the folder name under output\).
set /p WB="  Workbook name: "
if "%WB%"=="" goto menu
call Run_Verify.bat "%WB%" --open --refresh
goto menu

:validate
echo.
set /p PB="  Full path of the .pbip file: "
if "%PB%"=="" goto menu
call Run_Validate.bat "%PB%"
goto menu
