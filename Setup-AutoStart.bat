@echo off
setlocal
set "VBS=%~dp0Start-FollowupAddin-Hidden.vbs"

echo.
echo  Setting up auto-start for the Follow-up Add-in server...
echo  Task name: FollowupMeetingAddinServer
echo.

schtasks /create ^
  /tn "FollowupMeetingAddinServer" ^
  /tr "wscript.exe \"%VBS%\"" ^
  /sc ONLOGON ^
  /ru "%USERDOMAIN%\%USERNAME%" ^
  /rl HIGHEST ^
  /f >nul 2>&1

if %ERRORLEVEL% EQU 0 (
  echo  [OK]  Task created. The server will start silently at every login.
  echo.
  echo  To disable auto-start later:
  echo    schtasks /delete /tn "FollowupMeetingAddinServer" /f
  echo  Or open Task Scheduler and delete "FollowupMeetingAddinServer".
  echo.
  echo  Starting the server now for this session...
  wscript.exe "%VBS%"
  echo  Done. Wait a few seconds then reload the add-in in OWA.
) else (
  echo  [!!] Task Scheduler registration failed (may need admin rights).
  echo       Falling back: added shortcut to Windows Startup folder instead.
  set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
  copy /y "%VBS%" "%STARTUP%\FollowupAddinServer.vbs" >nul
  echo  [OK]  Shortcut added to Startup folder - will auto-start at next login.
  wscript.exe "%VBS%"
)

echo.
pause
