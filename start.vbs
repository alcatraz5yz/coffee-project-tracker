' PCS Dashboard – startet NUR den lokalen Server (versteckt, kein Fenster).
'
' Den Browser oeffnest du selbst: http://127.0.0.1:8090  (z.B. in Chrome).
' Laeuft der Server schon, passiert nichts (kein zweiter Prozess).
' Stoppen: stop.bat.

Option Explicit
Dim fso, sh, baseDir, url
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = baseDir
url = "http://127.0.0.1:8090"

' Laeuft der Server bereits? Dann nichts tun.
If ServerIsUp(url) Then WScript.Quit

' Server versteckt starten (kein Konsolenfenster).
sh.Run "node server.js", 0, False

' Antwortet der lokale Server auf der URL?
Function ServerIsUp(u)
  Dim h
  ServerIsUp = False
  On Error Resume Next
  Set h = CreateObject("MSXML2.XMLHTTP")
  h.open "GET", u, False
  h.send
  If Err.Number = 0 And h.status >= 200 And h.status < 600 Then ServerIsUp = True
  On Error GoTo 0
End Function
