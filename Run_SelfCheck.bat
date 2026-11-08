@echo off
rem Bhayankar self-check: tests the toolkit itself (regression tests in tests\). Run once on a new PC and after any change to the .py files
rem Optional corpus smoke test: set BHAYANKAR_CORPUS=<folder of .twb/.twbx> before running
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (
    python -m unittest discover -s tests -v %*
) else (
    py -3 -m unittest discover -s tests -v %*
)
set RC=%errorlevel%
echo.
if %RC%==0 (echo RESULT: SELF-CHECK PASSED - this PC gives the same results as everyone else) else (echo RESULT: SELF-CHECK FAILED - see output above, do not use the toolkit on this PC yet)
pause
exit /b %RC%
