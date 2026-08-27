' BPS Report Generator — Launcher (Hidden Window)
' Double-click ni untuk jalan app. Tiada terminal/CMD akan nampak.
' App akan buka http://localhost:5000 dalam browser.
' Nak matikan server, guna Stop BPS.vbs atau Task Manager.

Dim objShell, objFSO, appDir
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Dapatkan folder script ni berada
appDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Check kalau server dah jalan — elak double-start
Dim pid, objExec, line
pid = ""
On Error Resume Next
Set objExec = objShell.Exec("cmd /c netstat -ano | findstr :5000")
If Err.Number = 0 Then
    Do While Not objExec.StdOut.AtEndOfStream
        line = objExec.StdOut.ReadLine()
        If InStr(line, "LISTENING") > 0 Then
            Dim parts
            parts = Split(line)
            If UBound(parts) >= 4 Then
                pid = Trim(parts(UBound(parts)))
                Exit Do
            End If
        End If
    Loop
End If
On Error GoTo 0

If pid <> "" Then
    ' Server dah jalan — buka browser je
    objShell.Run "http://localhost:5000", 1, False
    WScript.Quit
End If

' Jalan pythonw.exe (tanpa tingkap)
objShell.Run """" & appDir & "\.venv\Scripts\pythonw.exe"" """ & appDir & "\app.py""", 0, False

' Tunggu 3 saat untuk server start, pastu buka browser
WScript.Sleep 3000
objShell.Run "http://localhost:5000", 1, False
