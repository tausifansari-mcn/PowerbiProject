@echo off
title GNC Overall Sale dashboard - auto create
echo ============================================================
echo  GNC Overall Sale Performance Dashboard - auto create
echo  Report date = yesterday (or type a date: GNC_Auto.bat 2026-10-08)
echo   1. Rohan's "GNC Dashboard Data" mail (Order + Allocation)
echo   2. CDR: GNC ^<day^> CDR from the morning outbound CDR, sorted by Priorty
echo   3. Sale, Zecpe, Allocation, APR, AZ2, pivots - dashboard saved
echo  Keep the dashboard and the CDR files CLOSED while it runs.
echo ============================================================
echo.
if "%~1"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0GNC_Auto.ps1"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0GNC_Auto.ps1" -ReportDate %1
)
echo.
if %errorlevel%==0 (
  echo  READY - see the DONE line above.
) else (
  echo  Something went wrong - see the lines above or ..\Logs\gnc-auto-log.txt
)
echo.
pause
