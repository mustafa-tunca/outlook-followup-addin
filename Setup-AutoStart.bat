@echo off
setlocal
set "VBS=%~dp0Start-FollowupAddin-Hidden.vbs"
set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "LNK=%STARTUP%\FollowupAddinServer.lnk"

echo.
echo  Setting up auto-start for the Follow-up Add-in server...
echo.

:: Remove any old file-copy from previous installs
if exist "%STARTUP%\FollowupAddinServer.vbs" del /f /q "%STARTUP%\FollowupAddinServer.vbs"

:: Create Startup folder shortcut pointing to the VBScript in the project folder
PowerShell -NonInteractive -Command "$s=New-Object -ComObject WScript.Shell; $l=$s.CreateShortcut('%LNK%'); $l.TargetPath='wscript.exe'; $l.Arguments='\"%VBS%\"'; $l.WorkingDirectory='%~dp0'; $l.WindowStyle=7; $l.Description='Follow-up Add-in dev server'; $l.Save()" >nul 2>&1

if exist "%LNK%" (
  echo  [OK]  Startup shortcut created - server will start silently at next login.
) else (
  echo  [!!]  Could not create Startup shortcut. You may need to run manually.
)

:: Also register Task Scheduler as a backup (20s delay so Windows PATH is ready)
PowerShell -NonInteractive -Command "$a=New-ScheduledTaskAction -Execute 'wscript.exe' -Argument '\"%VBS%\"'; $t=New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME; $t.Delay='PT20S'; $s=New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries; Register-ScheduledTask -TaskName 'FollowupMeetingAddinServer' -Action $a -Trigger $t -Settings $s -Force -EA SilentlyContinue" >nul 2>&1

echo  [OK]  Task Scheduler entry registered (backup, 20s delay after login).
echo.
echo  Starting server now for this session...
wscript.exe "%VBS%"
echo  Done. The add-in will be ready in about 15 seconds.
echo.
pause
