@echo off
setlocal
echo Running the automatic checks. No clicking needed.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-ExcelChecks.ps1" %*
echo.
echo Done. Results are in the "out" folder beside this file. Send me out\results.json
pause
