@echo off
rem Double-click to rerun only part 2 of the M3 Windows check (the --reopen script).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0m3_check.ps1" -ReopenOnly
echo.
pause
