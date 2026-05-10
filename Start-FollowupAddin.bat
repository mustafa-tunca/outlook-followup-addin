@echo off
title Follow-up Add-in - Dev Server
cd /d "%~dp0"

:: ── Check Node.js ──────────────────────────────────────────────────────────
where npm >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
  echo.
  echo  [ERROR] Node.js / npm not found in PATH.
  echo  Please install Node.js from https://nodejs.org (LTS version).
  echo  Then run this file again.
  echo.
  pause
  exit /b 1
)

:: ── Install dependencies if node_modules is missing ───────────────────────
if not exist "%~dp0node_modules" (
  echo  [INFO] node_modules not found - running npm install first...
  call npm install
  if %ERRORLEVEL% NEQ 0 (
    echo  [ERROR] npm install failed. Check your internet connection.
    pause
    exit /b 1
  )
)

echo.
echo  ============================================
echo   Follow-up Meeting Add-in ^| Dev Server
echo  ============================================
echo.
echo   Server  : https://localhost:3000
echo   Manifest: manifest.json  (OWA sideloading)
echo.
echo   Keep this window open while using
echo   the add-in in OWA. Close it to stop.
echo  ____________________________________________
echo.

call npm start

echo.
echo  Server stopped. Press any key to close.
pause >nul
