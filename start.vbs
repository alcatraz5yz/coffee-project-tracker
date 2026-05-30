' Launch the PCS Dashboard in the background (no console window).
' Usage: open a terminal in this folder and run `wscript start.vbs`.
' The server binds to 127.0.0.1 by default, so it is only reachable from this PC.
' To stop: run `stop.bat`.

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
sh.Run "node server.js", 0, False
