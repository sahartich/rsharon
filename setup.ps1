$originPath = $PSscriptroot
$destinationPath = ($env:psmodulepath).split(";")[1]

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

Move-Item -path "$originPath\rsharon.ps1" -Destination "$destinationPath\rsharon" -Force
Move-Item -path "$originPath\rdpsharon.ps1" -Destination "$destinationPath\rsharon" -Force
Remove-Item -path $PSCommandPath -Force