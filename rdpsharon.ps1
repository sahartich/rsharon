function rdpsharon{
    param(
    $computer
    )

    if ($null -eq $computer){
        $computer = Read-Host -Prompt "Enter the USER IP or PC NAME"
        if($computer -match '^[\d.]+$'){
            Write-Host "You've Entered the PC's IP Address"
        }
        else {
            Write-Host "You've Entered the PC's Name"
        }
    }

    $id = quser /server:$computer | ForEach-Object { $_ -replace '\s+', ' ' } | Select-Object -Skip 1 | ForEach-Object { $_.Split()[3] }
    try{
        mstsc /shadow:$id /v:$computer /control
    }
    catch{
        
    }
}

RDP_rsharon