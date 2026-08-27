' BPS Report Generator — Stop Server
' Hentikan server BPS yang sedang mendengar pada port 5000.
' Lebih selamat berbanding kill semua python.exe/pythonw.exe.

Dim objShell, objExec, line, pid, parts
Set objShell = CreateObject("WScript.Shell")

pid = ""
On Error Resume Next
Set objExec = objShell.Exec("cmd /c netstat -ano | findstr :5000")
If Err.Number = 0 Then
    Do While Not objExec.StdOut.AtEndOfStream
        line = objExec.StdOut.ReadLine()
        ' Format: TCP    127.0.0.1:5000    0.0.0.0:0    LISTENING    12345
        If InStr(line, "LISTENING") > 0 Then
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
    objShell.Run "taskkill /F /PID " & pid, 0, True
    MsgBox "✅ Server BPS (PID " & pid & ") telah dimatikan.", vbInformation, "BPS Report"
Else
    MsgBox "ℹ️ Tiada server BPS yang sedang berjalan pada port 5000.", vbInformation, "BPS Report"
End If
