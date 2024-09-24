Add-Type -AssemblyName System.Windows.Forms
$originPath = $PSscriptroot
$destinationPath = ($env:psmodulepath).split(";")[1]
if(-not ($destinationPath -eq "C:\Program Files\WindowsPowerShell\Modules")){
     [System.Windows.Forms.MessageBox]::Show("Installation unsuccessful, something wrong about your Powershell module directory, please Contact Sahar!", "Error - Module Directory Not Found")
    exit
}
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

if(-Not (Test-Path "$destinationPath\rsharon")){
    New-Item -Path "$destinationPath\rsharon" -ItemType Directory
}
if(-Not (Test-Path "$originPath\rsharon.txt")){
	[System.Windows.Forms.MessageBox]::Show("Installation files are missing! Please import them properly and try again!", "Error - Installation Files Not Found")
	exit
}

$readmeContent = @"
-------------------------------------------------
-----THIS CODE WAS WRITTEN BY SAHAR TICHOVER-----
-------------------------------------------------

TYPE RSHARON TO FIND THE COMPUTER NAME AND CONNECT TO IT
IF YOU KNOW THE NAME OF THE COMPUTER TYPE RDPSHARON INSTEAD

1. THE CODE UPDATES THE CSV FILES AUTOMATICALLY
2. QUERIES UPON THE STORED CSV FILES TO FIND THE RELEVANT COMPUTERS
3. CHECKS WHICH OF THOSE COMPUTERS ARE AVAILABLE FOR AN RDP CONNECTION
4. PROMPTS TO CHOOSE THE RELEVANT COMPUTER TO CONNECT TO

VERSION 1.3.3
CREATED SEVERAL FILES TO IMPROVE READABILITY AND BREAK DOWN THE CODE FURTHER FOR MORE OPTIMIZED AND FLEXIBLE USE
-----
VERSION 1.3.2
NOW STORES THE ACTIVEDIRECTORY INFORMATION WITHININ A USER-ENCRYPTED FOLDER AS A CSV FILE AND THEN QUERIES UPON THAT SAVED INFO
-----
VERSION 1.3.1
QUALITY OF LIFE CHANGES INCLUDING THE OPTION TO SEARCH AGAIN AT ANY STAGE AND EXIT WITHOUT CTRL+C AND ERROR HANDLING AND PREVENTION
-----
VERSION 1.3
NOW USES SYSTEM.NET.SOCKETS.TCPCLIENT INSTEAD OF TEST-NETCONNECTION; WHICH IS MORE BARE BONES BUT FAR FASTER, PERFECT FOR SIMPLY TESTING AN OPEN PORT
USING THIS .NET FUNCTION ALSO ENABLES USING ASYNCWAITHANDLE WHICH TERMINATES THE PROCESS IF IT TAKES MORE THAN A SECOND, REDUCING THE RUNTIME SIGNIFICANTLY
-----
VERSION 1.2
NOW CREATES A LIST OF ALL USERS AND ALL COMPUTERS WHEN FIRST RAN INSTEAD OF QUERYING ACTIVE-DIRECTORY EVERY TIME THE FUNCTION IS CALLED
-----
VERSION 1.1
NOW CHECKS IF EACH QUERIED PC ACCEPTS A TCP CONNECTION (AND THUS WILL SUPPORT THE RDP), IF NOT IT IS OMITTED FROM THE LIST
-----
"@


Set-content -path "$destinationPath\rsharon\README.md" -Value $readmeContent
Move-Item -path "$originPath\rsharon.txt" -Destination "$destinationPath\rsharon\rsharon.psm1" -Force
Start-Sleep -Seconds 3
Import-Module rsharon
Update-CsvJob
Remove-Item -path $PSCommandPath -Force