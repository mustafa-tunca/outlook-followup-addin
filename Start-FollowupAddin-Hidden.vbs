' Starts the Follow-up Add-in dev server silently at login.
' Waits 12 seconds for Windows to finish loading before starting.

WScript.Sleep 12000

Dim shell, fso, dir, bat
Set shell = WScript.CreateObject("WScript.Shell")
Set fso   = WScript.CreateObject("Scripting.FileSystemObject")

dir = fso.GetParentFolderName(WScript.ScriptFullName)
bat = dir & "\Start-FollowupAddin.bat"

If Not fso.FileExists(bat) Then
    WScript.Quit 1
End If

' Add Node.js to PATH for this process
Dim nodePath, currentPath
nodePath = "C:\Program Files\nodejs"
currentPath = shell.ExpandEnvironmentStrings("%PATH%")
If InStr(currentPath, "nodejs") = 0 Then
    If fso.FolderExists(nodePath) Then
        shell.Environment("Process")("PATH") = nodePath & ";" & currentPath
    End If
End If

' Run hidden (0 = no window, False = do not wait)
shell.Run "cmd.exe /c """ & bat & """", 0, False
