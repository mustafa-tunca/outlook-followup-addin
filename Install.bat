@echo off
setlocal
cd /d "%~dp0"
title Follow-up Add-in - First-Time Setup

echo.
echo  ================================================
echo   Follow-up Meeting Add-in - First-Time Setup
echo  ================================================
echo.

:: 1. Add Node.js to PATH
if exist "C:\Program Files\nodejs\npm.cmd" set "PATH=C:\Program Files\nodejs;%PATH%"
if exist "C:\Program Files (x86)\nodejs\npm.cmd" set "PATH=C:\Program Files (x86)\nodejs;%PATH%"

:: 2. Check Node.js
echo  [1/5] Checking Node.js...
where npm.cmd >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
  echo.
  echo  [ERROR] Node.js not found.
  echo.
  echo  Please install Node.js LTS from:
  echo    https://nodejs.org
  echo.
  echo  After installing, restart your PC and then reopen this file.
  pause
  exit /b 1
)
for /f "tokens=*" %%v in ('node --version') do set NODEVER=%%v
echo         Found Node.js %NODEVER% - OK
echo.

:: 3. Install npm packages
echo  [2/5] Installing npm packages (this may take a minute)...
call npm.cmd install
if %ERRORLEVEL% NEQ 0 (
  echo  [ERROR] npm install failed. Check your internet connection.
  pause & exit /b 1
)
echo         Done.
echo.

:: 4. Build the add-in
echo  [3/5] Building the add-in...
call npm.cmd run build
if %ERRORLEVEL% NEQ 0 (
  echo  [ERROR] Build failed. See output above.
  pause & exit /b 1
)
echo         Done.
echo.

:: 5. Install HTTPS dev certificates
echo  [4/5] Installing HTTPS dev certificates...
echo         (A UAC prompt may appear - click Yes to trust the certificate.)
echo.
call npx.cmd office-addin-dev-certs install
echo         Certificate setup complete.
echo.

:: 6. Register auto-start task (server starts silently at login)
echo  [5/5] Registering auto-start...
set "VBS=%~dp0Start-FollowupAddin-Hidden.vbs"
set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "LNK=%STARTUP%\FollowupAddinServer.lnk"

:: Remove any old file-copy from previous installs
if exist "%STARTUP%\FollowupAddinServer.vbs" del /f /q "%STARTUP%\FollowupAddinServer.vbs"

:: Create Startup folder shortcut pointing back to VBScript in project folder
PowerShell -NonInteractive -Command "$s=New-Object -ComObject WScript.Shell; $l=$s.CreateShortcut('%LNK%'); $l.TargetPath='wscript.exe'; $l.Arguments='\"%VBS%\"'; $l.WorkingDirectory='%~dp0'; $l.WindowStyle=7; $l.Description='Follow-up Add-in dev server'; $l.Save()" >nul 2>&1

:: Also register Task Scheduler as a backup (20s delay so Windows PATH is ready)
PowerShell -NonInteractive -Command "$a=New-ScheduledTaskAction -Execute 'wscript.exe' -Argument '\"%VBS%\"'; $t=New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME; $t.Delay='PT20S'; $s=New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries; Register-ScheduledTask -TaskName 'FollowupMeetingAddinServer' -Action $a -Trigger $t -Settings $s -Force -EA SilentlyContinue" >nul 2>&1

echo         Auto-start registered (Startup folder + Task Scheduler).
echo         The server will start silently after your next login.
echo.

:: Done
echo  ================================================
echo   Setup complete!
echo  ================================================
echo.
echo   NEXT STEP: Sideload the add-in in OWA
echo.
echo   1. Open Outlook on the Web (OWA)
echo   2. Open any calendar appointment
echo   3. Click "..." menu - "Get Add-ins" - "My Add-ins"
echo   4. Choose "Add from file..." and select:
echo      %~dp0manifest.json
echo   5. The "Create Follow-up" button will appear
echo      in the appointment ribbon.
echo.
echo   Starting the server now for this session...
echo.
wscript.exe "%VBS%"
echo   Server started silently in the background.
echo   It will be ready in about 15 seconds.
echo.
pause
