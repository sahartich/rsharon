Add-Type -AssemblyName System.Windows.Forms

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")){
    try{
        Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    }
    catch{
        [System.Windows.Forms.MessageBox]::Show("Admin priviledge is required to run this script!", "Error - Access Denied")
    }
    exit
}



Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\Microsoft.PowerShellScript.1\Shell\Open\Command" -Name "(Default)" -Value 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell_ise.exe "%1"'


$filePath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell_ise.exe"
$shortcutPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "PowerShell_ISE.lnk")

$wshShell = New-Object -ComObject WScript.Shell
$shortcut = $wshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $isePath
$shortcut.Save()