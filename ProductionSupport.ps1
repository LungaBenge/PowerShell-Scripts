# Create a Access Reports Shortcut with Windows PowerShell
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$Home\Desktop\Production Support.lnk")
$Shortcut.TargetPath = ""
$Shortcut.Save()