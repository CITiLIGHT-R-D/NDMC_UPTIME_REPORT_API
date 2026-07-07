@echo off
title NDMC Operational (Detailed) Report Generator
cd /d "%~dp0"

echo =========================================
echo    NDMC Operational (Detailed) Report
echo =========================================
echo.

REM --- Check Node.js is installed ---
where node >nul 2>nul
if errorlevel 1 (
  echo  ERROR: Node.js is not installed.
  echo  Please install it from https://nodejs.org and run this again.
  echo.
  pause
  exit /b 1
)

REM --- Install dependencies on first run only ---
if not exist "node_modules\exceljs" (
  echo  First-time setup: installing dependencies...
  call npm install
  echo.
)

REM --- Run the operational report (asks for month, username and password) ---
node NdmcOperationalReport.js

echo.
echo =========================================
if errorlevel 1 (
  echo  Something went wrong - see the messages above.
) else (
  echo  Done. The report has opened in Excel.
)
echo =========================================
echo.
pause
