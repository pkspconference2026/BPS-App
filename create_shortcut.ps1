$WS = New-Object -ComObject WScript.Shell
$SC = $WS.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\BPS Report.lnk')
$SC.TargetPath = 'C:\Users\User\Desktop\BPS-App\Jalan BPS.bat'
$SC.WorkingDirectory = 'C:\Users\User\Desktop\BPS-App'
$SC.Description = 'BPS Report Generator - Laporan Penilaian Biopsikososial'
$SC.IconLocation = 'C:\Users\User\Desktop\BPS-App\BPS.ico,0'
$SC.Save()
Write-Host 'Shortcut updated!'