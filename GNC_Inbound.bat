@echo off
title GNC Inbound MIS - daily raw data update
echo ============================================================
echo  GNC Inbound MIS Dashboard - daily raw data update
echo   1. Feedback    (192.168.11.244 CDR/CSAT Report, GNC)  -^> sheet "Feedback Raw"        formulas M:P dragged
echo   2. CDR Report IN (192.168.11.240)                     -^> sheet "CDR Raw"             formulas dragged
echo   3. In Call Details (DialDesk, GNC, All)               -^> sheet "Inbound DD Tagging"  formulas dragged
echo  Report date = yesterday (or type a date: GNC_Inbound.bat 2026-10-11)
echo  Missed days since the last date in each sheet are caught up automatically.
echo  Keep "GNC Inbound MIS Dashboard" CLOSED while it runs.
echo ============================================================
echo.
set "DATEARG="
set "WAITARG="
if /i "%~1"=="auto" set "WAITARG=-WaitMinutes 30"
if /i "%~2"=="auto" set "WAITARG=-WaitMinutes 30"
if not "%~1"=="" if /i not "%~1"=="auto" set "DATEARG=-ReportDate %~1"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0GNC_Inbound.ps1" %DATEARG% %WAITARG%
set "RC=%errorlevel%"
echo.
if "%RC%"=="0" (
  echo  READY - see the DONE line above.
) else (
  echo  Something went wrong - see the lines above or Logs\gnc-inbound-log.txt
)
echo.
if /i not "%~1"=="auto" if /i not "%~2"=="auto" pause
exit /b %RC%
