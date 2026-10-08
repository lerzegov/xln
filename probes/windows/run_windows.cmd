@echo off
rem Double-click to run the excel-dim name probe. The report lands in probes\results\.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0probe_names.ps1"
echo.
echo If the window above says the script cannot be loaded because running scripts is
echo disabled, open PowerShell, then copy and paste the whole of probe_names.ps1 into it.
pause
