' PCS Dashboard – startet den Server versteckt und oeffnet ein randloses
' App-Fenster in Microsoft Edge: KEINE Adressleiste, KEINE Tabs, kein Browser-
' Chrome. Wirkt wie ein eigenstaendiges Programm (passend zum Explorer-Theme).
'
' Doppelklick auf diese Datei genuegt. Stoppen: stop.bat.
'
' Warum Edge --app und kein Electron/Tauri:
'   Edge ist bereits installiert und IT-freigegeben. Kein neues Binary, keine
'   ~150 MB Chromium-Runtime, nichts das die EDR/Sicherheitssoftware triggert.

Option Explicit
Dim fso, sh, baseDir, url, edge, profileDir
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = baseDir
url = "http://127.0.0.1:8090"

' 1) Server versteckt starten (kein Konsolenfenster). Laeuft er schon, schadet es nicht.
sh.Run "node server.js", 0, False

' 2) Warten bis der Server antwortet (max. ~20 s), damit kein leeres Fenster aufgeht.
Dim http, ready, i
ready = False
For i = 1 To 40
  On Error Resume Next
  Set http = CreateObject("MSXML2.XMLHTTP")
  http.open "GET", url, False
  http.send
  If Err.Number = 0 And http.status >= 200 And http.status < 500 Then ready = True
  On Error GoTo 0
  If ready Then Exit For
  WScript.Sleep 500
Next

' 3) Edge suchen.
edge = FindEdge(fso, sh)
If edge = "" Then
  ' Fallback: im Standardbrowser oeffnen (immerhin nutzbar) und melden.
  sh.Run """" & url & """", 1, False
  MsgBox "Microsoft Edge wurde nicht gefunden – das Dashboard wurde im " & _
         "Standardbrowser geoeffnet. Fuer den randlosen App-Modus bitte Edge " & _
         "installieren lassen.", vbInformation, "PCS Dashboard"
  WScript.Quit
End If

' 4) Eigenes, isoliertes Profil -> sauberes Fenster ohne fremde Tabs; merkt sich
'    Groesse und Position zwischen den Starts.
profileDir = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\PcsDashboardApp"

' 5) Edge im App-Modus oeffnen (randlos, ohne Adressleiste/Tabs) UND warten,
'    bis das Fenster geschlossen wird (letzter Parameter True = blockierend).
'    Dank eigenem --user-data-dir ist dieses Edge-Fenster ein eigener Prozess,
'    der erst beim Schliessen endet.
sh.Run """" & edge & """ --app=" & url & _
       " --user-data-dir=""" & profileDir & """" & _
       " --no-first-run --no-default-browser-check", 1, True

' 6) Fenster geschlossen -> Server sauber beenden. Kein lingernder, versteckter
'    node.exe-Prozess (waere fuer eine EDR-Verhaltensbasis eher auffaellig).
On Error Resume Next
sh.Run "cmd /c """ & baseDir & "\stop.bat""", 0, True
On Error GoTo 0
WScript.Quit

' --- Edge an gaengigen Orten + Registry finden -------------------------------
Function FindEdge(fsoRef, shRef)
  Dim cands, c, p
  cands = Array( _
    shRef.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Microsoft\Edge\Application\msedge.exe", _
    shRef.ExpandEnvironmentStrings("%ProgramFiles%")      & "\Microsoft\Edge\Application\msedge.exe", _
    shRef.ExpandEnvironmentStrings("%LOCALAPPDATA%")      & "\Microsoft\Edge\Application\msedge.exe")
  FindEdge = ""
  For Each c In cands
    If fsoRef.FileExists(c) Then
      FindEdge = c
      Exit Function
    End If
  Next
  ' Registry: App Paths
  On Error Resume Next
  p = shRef.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\")
  If fsoRef.FileExists(p) Then FindEdge = p
  On Error GoTo 0
End Function
