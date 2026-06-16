' Legt eine Desktop-Verknuepfung "PCS Dashboard" an, die das Dashboard im
' randlosen Edge-App-Modus startet (ueber start.vbs). Als Symbol wird das ECHTE
' Windows-Datei-Explorer-Icon verwendet (explorer.exe,0) -> auf dem Desktop nicht
' vom echten Explorer zu unterscheiden.
'
' Einmal doppelklicken. Danach das Desktop-Icon nutzen.

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
lnk.Description       = "PCS Kaffee Dashboard"
' Echtes Datei-Explorer-Icon aus Windows.
lnk.IconLocation     = sh.ExpandEnvironmentStrings("%SystemRoot%") & "\explorer.exe,0"
lnk.Save

MsgBox "Verknuepfung 'PCS Dashboard' wurde auf dem Desktop angelegt.", _
       vbInformation, "PCS Dashboard"
