@echo off
setlocal
echo Running the full checks, including the two macros that open a dialog.
echo The script will stop and tell you what to click.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-ExcelChecks.ps1" -Interactive %*
echo.
echo Done. Results are in the "out" folder beside this file. Send me out\results.json
pause
