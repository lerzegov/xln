@echo off
rem Double-click to run the xln M3 Windows check. The report lands in probes\results\,
rem Excel's saved copies in probes\results\m3win\. See probes\windows\M3-CHECKS.md.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0m3_check.ps1"
echo.
echo If the window above says the script cannot be loaded because running scripts is
echo disabled, open PowerShell in this folder and run:
echo     powershell -NoProfile -ExecutionPolicy Bypass -File .\m3_check.ps1
pause
