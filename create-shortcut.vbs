' Legt eine Desktop-Verknuepfung "PCS Dashboard" an, die den lokalen Server
' startet (ueber start.vbs, versteckt). Den Browser oeffnest du danach selbst:
' http://127.0.0.1:8090  (z.B. in Chrome).
'
' Einmal doppelklicken. Danach das Desktop-Icon zum Starten nutzen.

Option Explicit
Dim fso, sh, baseDir, desktop, lnk, target

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
desktop = sh.SpecialFolders("Desktop")
target  = baseDir & "\start.vbs"

Set lnk = sh.CreateShortcut(desktop & "\PCS Dashboard.lnk")
' wscript startet start.vbs ohne Konsolenfenster.
lnk.TargetPath       = "wscript.exe"
lnk.Arguments        = """" & target & """"
lnk.WorkingDirectory = baseDir
lnk.WindowStyle      = 1
lnk.Description       = "PCS Server starten – danach http://127.0.0.1:8090 im Browser oeffnen"
lnk.IconLocation     = sh.ExpandEnvironmentStrings("%SystemRoot%") & "\explorer.exe,0"
lnk.Save

MsgBox "Verknuepfung 'PCS Dashboard' wurde auf dem Desktop angelegt." & vbCrLf & _
       "Sie startet den Server – oeffne danach http://127.0.0.1:8090 selbst im Browser.", _
       vbInformation, "PCS Dashboard"
