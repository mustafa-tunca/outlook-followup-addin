' Starts the Follow-up Add-in dev server silently (no visible window).
' Called by the Windows Task Scheduler task at every login.
Dim shell, fso, dir, bat
Set shell = WScript.CreateObject("WScript.Shell")
Set fso   = WScript.CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
bat = dir & "\Start-FollowupAddin.bat"
' 0 = hidden window, False = fire-and-forget (don't block)
shell.Run "cmd /c """ & bat & """", 0, False
