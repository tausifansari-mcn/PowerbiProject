@echo off
title GNC Chat MIS - daily chat data update
echo ============================================================
echo  GNC Chat Dashboard - daily chat data update
echo   1. KwikEngage (app.kwikengage.ai) Analytics - Export Reports - Ticket
echo   2. link from the mail  -^> chat csv download
echo   3. sheet "Chat": new tickets appended + formulas dragged, changed tickets updated
echo   4. sheet "Snap": days shown up to yesterday
echo  Report date = yesterday (or type a date: GNCchatmis.bat 2026-10-11)
echo  Keep "GNC_Chat_Dashboard" CLOSED while it runs. The mail can take a few minutes.
echo ============================================================
echo.
set "DATEARG="
set "WAITARG="
if /i "%~1"=="auto" set "WAITARG=-WaitMinutes 30"
if /i "%~2"=="auto" set "WAITARG=-WaitMinutes 30"
if not "%~1"=="" if /i not "%~1"=="auto" set "DATEARG=-ReportDate %~1"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0GNCchatmis.ps1" %DATEARG% %WAITARG%
set "RC=%errorlevel%"
echo.
if "%RC%"=="0" (
  echo  READY - see the DONE line above.
) else (
  echo  Something went wrong - see the lines above or Logs\gnc-chatmis-log.txt
)
echo.
if /i not "%~1"=="auto" if /i not "%~2"=="auto" pause
exit /b %RC%
