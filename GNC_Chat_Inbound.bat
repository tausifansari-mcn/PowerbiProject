@echo off
title GNC Chat + Inbound sale sync
echo ============================================================
echo  GNC Chat + Inbound sale sync
echo  From "Date ^& Camp wise Overall Sale" (Overall Sale dashboard):
echo   - Campaign = Chat    -^> GNC_Chat_Dashboard      sheet "Sale Raw"  (A:S)
echo   - Campaign = Inbound -^> GNC Inbound MIS Dashboard sheet "Sales Raw" (A:N)
echo  New orders are appended below, wrong rows corrected, nothing duplicated.
echo  Month = month of yesterday (or type a date: GNC_Chat_Inbound.bat 2026-10-08)
echo  Run it after GNC_Auto.bat. Keep the Chat and Inbound dashboards CLOSED.
echo ============================================================
echo.
if "%~1"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0GNC_Chat_Inbound.ps1"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0GNC_Chat_Inbound.ps1" -ReportDate %1
)
echo.
if %errorlevel%==0 (
  echo  READY - see the DONE line above.
) else (
  echo  Something went wrong - see the lines above or Logs\gnc-chat-inbound-log.txt
)
echo.
pause
