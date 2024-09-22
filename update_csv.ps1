$global:path = "C:\rsharon_csv"

function Update-CSV{
    Import-Module ActiveDirectory
    if (-Not (Test-Path -Path $global:path)) {
        New-Item -ItemType Directory -Path $global:path
        cipher /e /s:$global:path
    }
    $users = Get-ADUser -Filter {Enabled -eq $True} -Properties mailNickname, GivenName, DisplayName, CN, MobilePhone, emailAddress, SamAccountName
    $users | Select-Object mailNickname, GivenName, DisplayName, CN, MobilePhone, emailAddress,SamAccountName |
        Export-CSV -path "$global:path\user_db.csv" -NoTypeInformation -Encoding UTF8 -Force
    $computers = Get-ADComputer -Filter *
    $computers | Select-Object Name |
        Export-CSV -path "$global:path\computer_db.csv" -NoTypeInformation -Encoding UTF8 -Force

}

Update-CSV