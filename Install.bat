@echo off
setlocal
cd /d "%~dp0"
title Follow-up Add-in - First-Time Setup

echo.
echo  ================================================
echo   Follow-up Meeting Add-in ^| First-Time Setup
echo  ================================================
echo.

:: ── 1. Check Node.js ───────────────────────────────────────────────────────
echo  [1/5] Checking Node.js...
where npm >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
  echo.
  echo  [ERROR] Node.js not found.
  echo.
  echo  Please install Node.js LTS from:
  echo    https://nodejs.org
  echo.
  echo  After installing, reopen this file.
  pause
  exit /b 1
)
for /f "tokens=*" %%v in ('node --version') do set NODEVER=%%v
echo          Found Node.js %NODEVER% - OK
echo.

:: ── 2. Install npm packages ────────────────────────────────────────────────
echo  [2/5] Installing npm packages...
call npm install
if %ERRORLEVEL% NEQ 0 (
  echo  [ERROR] npm install failed. Check your internet connection.
  pause & exit /b 1
)
echo          Done.
echo.

:: ── 3. Build the add-in ────────────────────────────────────────────────────
echo  [3/5] Building the add-in...
call npm run build
if %ERRORLEVEL% NEQ 0 (
  echo  [ERROR] Build failed. See output above.
  pause & exit /b 1
)
echo          Done.
echo.

:: ── 4. Install HTTPS dev certificates ─────────────────────────────────────
echo  [4/5] Installing HTTPS dev certificates...
echo         (A UAC prompt may appear - click Yes to trust the certificate.)
echo.
call npx office-addin-dev-certs install
echo          Certificate setup complete.
echo.

:: ── 5. Register auto-start task ────────────────────────────────────────────
echo  [5/5] Registering auto-start (server will launch silently at login)...
set "VBS=%~dp0Start-FollowupAddin-Hidden.vbs"

PowerShell -NonInteractive -Command ^
  "$a=New-ScheduledTaskAction -Execute 'wscript.exe' -Argument '\"%VBS%\"';" ^
  "$t=New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME;" ^
  "$s=New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries;" ^
  "Register-ScheduledTask -TaskName 'FollowupMeetingAddinServer' -Action $a -Trigger $t -Settings $s -Force -EA Stop;" >nul 2>&1

if %ERRORLEVEL% EQU 0 (
  echo          Auto-start task registered - OK
) else (
  echo          Task Scheduler failed (possibly restricted by IT policy).
  echo          Falling back: copying to Windows Startup folder...
  set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
  copy /y "%VBS%" "%STARTUP%\FollowupAddinServer.vbs" >nul
  if %ERRORLEVEL% EQU 0 (
    echo          Startup folder shortcut created - OK
  ) else (
    echo          [WARN] Could not register auto-start.
    echo          You can manually start the server by running:
    echo            Start-FollowupAddin.bat
  )
)
echo.

:: ── Done ───────────────────────────────────────────────────────────────────
echo  ================================================
echo   Setup complete!
echo  ================================================
echo.
echo   NEXT STEP: Sideload the add-in in OWA
echo.
echo   1. Open Outlook on the Web (OWA)
echo   2. Open any calendar appointment
echo   3. Click "..." menu ^> "Get Add-ins" ^> "My Add-ins"
echo   4. Choose "Add from file..." and select:
echo      %~dp0manifest.json
echo   5. The "Create Follow-up" button will appear
echo      in the appointment ribbon.
echo.
echo   Starting the server now for this session...
echo.
wscript.exe "%VBS%"
echo   Server started in background - OK
echo.
pause
