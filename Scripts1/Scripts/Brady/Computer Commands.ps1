#===================================================================
Function SendMessage
{
    REG ADD "\\$Computer\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v AllowRemoteRPC /t REG_DWORD /d 1 /f
    $Message = Read-Host "Message"
    $SendMsg = MSG console /Server:$Computer /Time:6000 $Message
}

#===================================================================
Function LockComputer
{
    $Lock = TSDiscon Console /Server:$Computer
}

#===================================================================
Function LogoffUser
{
    $Logoff = Reset Session Console /Server:$Computer
}

#===================================================================
Function RebootComputer
{
    Write-Host " "
    Write-Host "1 - 1 Minute"
    Write-Host "2 - 10 Minutes"
    Write-Host "3 - 30 Minutes"
    Write-Host "4 - 1 Hour"
    Write-Host "5 - Now"
    $Ans = Read-Host "Make Selection"
    
    If ($Ans -eq 1)
    {
        Write-Host
        $Time = 60
    }
    If ($Ans -eq 2)
    {
        Write-Host
        $Time = 600
    }
    If ($Ans -eq 3)
    {
        Write-Host
        $Time = 1800
    }
    If ($Ans -eq 4)
    {
        Write-Host
        $Time = 3600
    }
    If ($Ans -eq 5)
    {
        Write-Host
        $Time = 0
    }
    If ($Time -gt 0)
    {
        $RTime = $Time/60
    }
    $Input = Read-Host "Comment"
    $Comment = $Input+": You will be rebooted in $RTime minute(s), please save all work"
    $Shutdown = Shutdown /r /f /m \\$Computer /t $Time /c $Comment
}

#===================================================================

Do
{
    Cls
    Write-Host " "
    Write-Host "1 - Send Message"
    Write-Host "2 - Lock Computer"
    Write-Host "3 - Logoff User"
    Write-Host "4 - Reboot Computer"
    Write-Host "5 - Exit"
    Write-Host " "

    $Ans = Read-Host "Make Selection"
    
    If ($Ans -eq 1)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        SendMessage
    }
    If ($Ans -eq 2)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        LockComputer
    }
    If ($Ans -eq 3)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        LogoffUser
    }
    If ($Ans -eq 4)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        RebootComputer
    }
}
Until ($Ans -eq 5)