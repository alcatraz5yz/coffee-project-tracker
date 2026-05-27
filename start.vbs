' Launch the PCS Dashboard in the background (no console window).
' Usage: open a terminal in this folder and run `wscript start.vbs`.
' Closing the terminal afterwards does not stop the server.
' To stop: Task Manager → end node.exe, or `taskkill /IM node.exe /F`.

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
sh.Run "node server.js", 0, False
