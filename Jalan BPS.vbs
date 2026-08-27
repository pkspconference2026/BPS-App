' BPS Report Generator — Launcher (Hidden Window)
' Double-click ni untuk jalan app. Tiada terminal/CMD akan nampak.
' App akan buka http://localhost:5000 dalam browser.
' Nak matikan server, guna Stop BPS.vbs atau Task Manager.

Dim objShell, objFSO, appDir
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Dapatkan folder script ni berada
appDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Jalan pythonw.exe (tanpa tingkap)
objShell.Run "\"" & appDir & "\.venv\Scripts\pythonw.exe\" \"" & appDir & "\app.py\"", 0, False

' Tunggu 3 saat untuk server start, pastu buka browser
WScript.Sleep 3000
objShell.Run "http://localhost:5000", 1, False
