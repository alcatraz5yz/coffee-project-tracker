' PCS Dashboard – startet den Server versteckt und oeffnet ein randloses
' App-Fenster in Microsoft Edge: KEINE Adressleiste, KEINE Tabs, kein Browser-
' Chrome. Wirkt wie ein eigenstaendiges Programm (passend zum Explorer-Theme).
'
' Doppelklick auf diese Datei genuegt. Beim Schliessen des Fensters wird der
' Server sauber beendet (kein lingernder node.exe).
'
' Single-Instance-Schutz: Nur die Instanz, die den Server tatsaechlich startet
' (Port war frei), beendet ihn beim Schliessen. Ein versehentlicher Zweitklick
' findet den Server bereits laufend, tut NICHTS und kann ihn damit nicht
' "unter dir wegkillen".
'
' Warum Edge --app und kein Electron/Tauri:
'   Edge ist bereits installiert und IT-freigegeben. Kein neues Binary, keine
'   ~150 MB Chromium-Runtime, nichts das die EDR/Sicherheitssoftware triggert.

Option Explicit
Dim fso, sh, baseDir, url, edge, profileDir, lockFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = baseDir
url = "http://127.0.0.1:8090"
lockFile = baseDir & "\.app-running.lock"

' === Single-Instance-Schutz =================================================
' 1) Antwortet der Server bereits? Dann ist schon ein Fenster offen -> diese
'    Zweitinstanz startet nichts, killt nichts und endet sofort.
If ServerIsUp(url) Then WScript.Quit

' 2) Sub-Sekunden-Doppelklick: ein FRISCHES Lock heisst, eine andere Instanz ist
'    gerade im Hochfahren (Port noch nicht offen) -> ebenfalls beenden.
'    Ein altes Lock (z.B. nach einem Absturz) gilt als veraltet und wird ignoriert.
If LockIsFresh(fso, lockFile) Then WScript.Quit

' Wir sind die Primaerinstanz und uebernehmen den Server-Lebenszyklus.
WriteLock fso, lockFile

' === Start ==================================================================
' 3) Server versteckt starten (kein Konsolenfenster).
sh.Run "node server.js", 0, False

' 4) Warten bis der Server antwortet (max. ~20 s), damit kein leeres Fenster aufgeht.
Dim i
For i = 1 To 40
  If ServerIsUp(url) Then Exit For
  WScript.Sleep 500
Next

' 5) Edge suchen.
edge = FindEdge(fso, sh)
If edge = "" Then
  ' Fallback: im Standardbrowser oeffnen. Da wir das Fenster dann nicht
  ' ueberwachen koennen, laeuft der Server weiter (Stop ueber stop.bat).
  RemoveLock fso, lockFile
  sh.Run """" & url & """", 1, False
  MsgBox "Microsoft Edge wurde nicht gefunden – das Dashboard wurde im " & _
         "Standardbrowser geoeffnet. Fuer den randlosen App-Modus bitte Edge " & _
         "installieren lassen. (Stoppen: stop.bat)", vbInformation, "PCS Dashboard"
  WScript.Quit
End If

' 6) Eigenes, isoliertes Profil -> sauberes Fenster ohne fremde Tabs; merkt sich
'    Groesse und Position zwischen den Starts.
profileDir = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\PcsDashboardApp"

' 7) Edge im App-Modus oeffnen (randlos) UND warten, bis das Fenster geschlossen
'    wird (letzter Parameter True = blockierend). Dank eigenem --user-data-dir ist
'    dieses Edge-Fenster ein eigener Prozess, der erst beim Schliessen endet.
sh.Run """" & edge & """ --app=" & url & _
       " --user-data-dir=""" & profileDir & """" & _
       " --no-first-run --no-default-browser-check", 1, True

' 8) Fenster geschlossen -> Server gezielt ueber die Port-PID beenden (stop.bat)
'    und Lock entfernen. Kein lingernder, versteckter node.exe-Prozess.
On Error Resume Next
sh.Run "cmd /c """ & baseDir & "\stop.bat""", 0, True
On Error GoTo 0
RemoveLock fso, lockFile
WScript.Quit

' === Hilfsfunktionen ========================================================

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

' Existiert ein Lock, das juenger als 30 s ist? (deckt das Startfenster ab)
Function LockIsFresh(fsoRef, path)
  Dim f
  LockIsFresh = False
  On Error Resume Next
  If fsoRef.FileExists(path) Then
    Set f = fsoRef.GetFile(path)
    If DateDiff("s", f.DateLastModified, Now) < 30 Then LockIsFresh = True
  End If
  On Error GoTo 0
End Function

Sub WriteLock(fsoRef, path)
  Dim t
  On Error Resume Next
  Set t = fsoRef.CreateTextFile(path, True)
  t.WriteLine "running " & Now
  t.Close
  On Error GoTo 0
End Sub

Sub RemoveLock(fsoRef, path)
  On Error Resume Next
  If fsoRef.FileExists(path) Then fsoRef.DeleteFile path, True
  On Error GoTo 0
End Sub

' Edge an gaengigen Orten + Registry finden.
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
  On Error Resume Next
  p = shRef.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\")
  If fsoRef.FileExists(p) Then FindEdge = p
  On Error GoTo 0
End Function
