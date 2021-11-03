Function Magic-Number
{
    Param (
        [Parameter(Mandatory=$true)]
        [Int] 
        $Systems
    ,
        [Parameter(Mandatory=$true)]
        [Int]
        $CAT1s
    ,
        [Parameter(Mandatory=$true)]
        [Int]
        $CAT2s
    ,
        [Parameter(Mandatory=$true)]
        [Int]
        $CAT3s
        )
    $MagicNumber = [Math]::Round((($CAT1s*10)+($CAT2s*4)+$CAT3s)/($Systems*15), 2)

    If ($MagicNumber -lt 2.5)
    {
        Write-Host
        Write-Host "Magic Number:" -NoNewline "$MagicNumber" -ForegroundColor Green
    }
    If ($MagicNumber -lt 5 -and $MagicNumber -gt 2.5)
    {
        Write-Host
        Write-Host "Magic Number:" -NoNewline "$MagicNumber" -ForegroundColor Yellow
    }
    If ($MagicNumber -gt 5)
    {
        Write-Host
        Write-Host "Magic Number:" -NoNewline "$MagicNumber" -ForegroundColor Red
    }
}

Do
{
    Cls
    Write-Host
    Write-Host "1 - Magic Number"
    Write-Host "2 - Exit"
    Write-Host

    $Ans = Read-Host "Make Selection"
    
    If ($Ans -eq 1)
    {
        Cls  
        Magic-Number
        Write-Host
        Pause
    }
}
Until ($Ans -eq 2)