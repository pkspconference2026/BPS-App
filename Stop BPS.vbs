' BPS Report Generator — Stop Server
' Double-click ni untuk matikan server BPS yang sedang berjalan.
' Alternatif: Buka Task Manager → cari python.exe / pythonw.exe → End Task.

Dim objShell, strCmd
Set objShell = CreateObject("WScript.Shell")

' Kill semua proses python yang jalan app.py dari folder BPS-App
strCmd = "taskkill /F /IM pythonw.exe 2>nul"
objShell.Run "cmd /c " & strCmd, 0, True

' Juga kill python.exe kalau ada (versi lama)
objShell.Run "cmd /c taskkill /F /IM python.exe 2>nul", 0, True

MsgBox "✅ Server BPS dah dimatikan.", vbInformation, "BPS Report"