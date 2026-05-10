' Starts the Follow-up Add-in dev server silently at login.
' Uses a short delay so Windows finishes loading the user environment first.
' Explicitly adds Node.js to PATH so npm is always found.

WScript.Sleep 8000   ' 8 s — enough for Node.js PATH to be available

Dim shell, fso, dir, bat, nodePath, currentPath
Set shell = WScript.CreateObject("WScript.Shell")
Set fso   = WScript.CreateObject("Scripting.FileSystemObject")

dir = fso.GetParentFolderName(WScript.ScriptFullName)
bat = dir & "\Start-FollowupAddin.bat"

' Ensure Node.js is in PATH — covers both system-wide and per-user installs
nodePath = "C:\Program Files\nodejs"
currentPath = shell.ExpandEnvironmentStrings("%PATH%")
If InStr(currentPath, "nodejs") = 0 Then
    shell.Environment("Process")("PATH") = nodePath & ";" & currentPath
End If

' Run the batch file hidden (0 = no window, False = fire-and-forget)
shell.Run """" & bat & """", 0, False
