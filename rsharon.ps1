<#
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
THIS CODE WAS WRITTEN BY SAHAR TICHOVER
IT IS STILL UNDER DEVELOPMENT, SOME FEATURES ARE MISSING

TYPE EITHER "RDPC" IF YOU KNOW THE PC'S IP/NAME
OR "SUGCC" OTHERWISE

PLEASE ONLY OPEN THIS FILE WITH POWERSHELL ISE AS POWERSHELL/CMD DO NOT SUPPORT DISPLAYING HEBREW WITH ACTIVE DIRECTORY QUERIES SPECIFICALLY
#>

Import-Module ActiveDirectory
$global:users = Get-ADUser -Filter {Enabled -eq $True} -Properties mailNickname, GivenName, DisplayName, CN, MobilePhone, emailAddress
$global:computers = Get-ADComputer -Filter *

function Select-ValidNumber{
    param(
    [string]$prompt,
    [int]$range,
    [bool]$searchAgain = $false
    )

    $hostInput = Read-Host -Prompt "$prompt or exit to abort"
    if ($hostInput -eq "exit"){
        return $hostInput
    }
    elseif(-not ($hostInput -match '^\d+$')){
        Write-Host "Not a number" -ForegroundColor Yellow
        Select-ValidNumber -prompt $prompt -range $range -searchAgain $searchAgain
    }
    elseif($searchAgain -and [int]$hostInput -eq 0){
        Write-Host "starting over........" -ForegroundColor Green
        return $hostInput
    }
    elseif(-not([int]$hostInput -gt 0 -and [int]$hostInput -lt $range)){  # less than and not less equal because counter always returns one above the highest number in the array
        Write-Host "Not a valid number" -ForegroundColor Yellow
        Select-ValidNumber -prompt $prompt -range $range -searchAgain $searchAgain
    }
    else{
        return $hostInput
    }
}

function rsharon{
    while($True){
        $userInput = Read-Host -Prompt "Enter the user's name [Hebrew or English]"
        $userInput = "*"+$userInput+"*"
        $userList = $global:users | Where-Object {$_.Name -like $userInput -or $_.SamAccountName -like $userInput -or $_.mailNickname -like $userInput -or $_.GivenName -like $userInput -or $_.DisplayName -like $userInput -or $_.CN -like $userInput}
        $userCounter=1
        $userArray = @()

        foreach ($user in $userList) {
            $userArrayObject = New-Object PSObject         
            $userArrayObject | Add-Member -MemberType NoteProperty -Name "Number" -Value $userCounter
            $userArrayObject | Add-Member -MemberType NoteProperty -Name "Username" -Value $user.SamAccountName
            $userArrayObject | Add-Member -MemberType NoteProperty -Name "Phone Number" -Value $user.MobilePhone
            $userArrayObject | Add-Member -MemberType NoteProperty -Name "Email Address" -Value $user.emailAddress
            $userArrayObject | Add-Member -MemberType NoteProperty -Name "Full Name" -Value $user.DisplayName
            $userArray += $userArrayObject
            $userCounter++
        }
        if($userArray.count -eq 0){
            Write-Host "No users found, please try again......." -ForegroundColor Yellow
            continue
        }

        $userArray | Format-Table | Out-String | ForEach-Object {Write-Host $_}
        $userInput = Select-ValidNumber -Prompt "Select a user by typing a number, 0 to search again" -range $userCounter -searchAgain $True
        if($userInput -eq "exit"){ return }
        elseif($userInput -eq 0){ continue }
        break
    }

    $user = ($userArray | Where-Object {$_.Number -eq $userInput} | Select-Object Username).Username
    $user1 = "*" + $user.split("_")[0] + "_PC*"
    $user2 = "*" + $user.split("_")[0] + "_NEW*"
    $user3 = "*" + $user.split("_")[0] + "_LAP*"
    $userDash =  "*" + $user.split("_")[0] + "-*"
    $userConnected = $user -replace "_", ""
    $userConnectedUnderscore = "*" + $userConnected +"_*"
    $userConnectedDash = "*" + $userConnected +"-*"
    $user = "*" + $user + "*"
    $computerList = $global:computers | Where-Object {$_.Name -like $user -or $_.Name -like $user1 -or $_.Name -like $user2 -or $_.Name -like $user3 -or $_.Name -like $userDash -or $_.Name -like $userConnectedUnderscore -or $_.Name -like $userConnectedDash}

    $computerArray = @()
    $computerCounter=1
    foreach ($computer in $computerList) {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        try{
            $result = $tcpClient.BeginConnect($computer.Name,3389,$null,$null)
            $success = $result.AsyncWaitHandle.WaitOne(1000,$false)
            if($success){
                $tcpClient.EndConnect($result)
                $computerArrayObject = New-Object PSObject
                $computerArrayObject | Add-Member -MemberType NoteProperty -Name "Number" -Value $computerCounter
                $computerArrayObject | Add-Member -MemberType NoteProperty -Name "Name" -Value $computer.Name
                $computerArray += $computerArrayObject
                $computerCounter++
            }
            else{Write-Host "Couldn't establish connection to $($computer.Name), it's either offline or RDP is blocked..." -ForegroundColor Yellow}
         }
         catch {Write-Host "AN UNEXPECTED ERROR WAS FOUND $($_.Exception.Message), PLEASE REPORT THIS TO SAHAR IMMEDIATELY" -ForegroundColor Red}
    }

    if($computerArray.Count -eq 0){
        Write-Host "No Available Computers Found, try Bomgar..." -ForegroundColor Red
        return
    }

    $computerArray | Format-Table | Out-String |ForEach-Object {Write-Host $_}
    $computerInput = Select-ValidNumber -prompt "Select a computer by typing a number" -range $computerCounter
    if($computerInput -eq "exit"){
        return
    }
    $selectedComputer = ($computerArray | Where-Object {$_.Number -eq $computerInput} | Select-Object Name).Name

    $id = quser /server:$selectedComputer | ForEach-Object { $_ -replace '\s+', ' ' } | Select-Object -Skip 1 | ForEach-Object { $_.Split()[3] }
    mstsc /shadow:$id /v:$selectedComputer /control
}