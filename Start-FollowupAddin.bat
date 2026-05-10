@echo off
title Follow-up Add-in - Dev Server
cd /d "%~dp0"

:: Add Node.js to PATH (handles both system and per-user installs)
if exist "C:\Program Files\nodejs\npm.cmd" set "PATH=C:\Program Files\nodejs;%PATH%"
if exist "C:\Program Files (x86)\nodejs\npm.cmd" set "PATH=C:\Program Files (x86)\nodejs;%PATH%"

:: Verify npm is available
where npm.cmd >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
  echo.
  echo [ERROR] npm.cmd not found. Install Node.js from https://nodejs.org
  echo.
  pause
  exit /b 1
)

:: Install dependencies if node_modules is missing
if not exist "%~dp0node_modules\" (
  echo [INFO] node_modules not found - running npm install first...
  call npm.cmd install
  if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] npm install failed.
    pause
    exit /b 1
  )
)

echo.
echo  ============================================
echo   Follow-up Meeting Add-in - Dev Server
echo  ============================================
echo.
echo   Server  : https://localhost:3000
echo   Manifest: manifest.json
echo.
echo   Keep this window open while using
echo   the add-in in OWA. Close it to stop.
echo  ____________________________________________
echo.

call npm.cmd start

echo.
echo  Server stopped. Press any key to close.
pause >nul
