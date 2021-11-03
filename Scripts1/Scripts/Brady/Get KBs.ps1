#========================================================================================
Function Get-KB
{
    $Array = @()
    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            $Comp = (Get-WmiObject Win32_ComputerSystem -cn $Computer).Name
            $KBs = Get-HotFix -cn $Comp
            ForEach ($KB in $KBs)
            {
                $obj = New-Object PSObject
                $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Comp
                $obj | Add-Member -Force -MemberType NoteProperty -Name "KBNumber" -Value $KB.HotFixId
                $obj | Add-Member -Force -MemberType NoteProperty -Name "InstalledOn" -Value $KB.InstalledOn
                $obj | Add-Member -Force -MemberType NoteProperty -Name "Description" -Value $KB.Description
                $Array += $obj
            }
        }
        $Count = ($Array | Where-Object {$_.ComputerName -like $Comp} | measure).count
        Write-Host "$Comp has $Count MS patches installed" -ForegroundColor Yellow
        Write-Host
    }
    $Array | Where-Object {$_.KBNumber -like "*$Search*"} | OGV -Title "$Comp MS patches installed"
    Pause
}

#========================================================================================
Do
{
    $Search = ""
    Cls
    Write-Host
    Write-Host "1 - All MS Patches Single Computer"
    Write-Host "2 - Specific MS Patch Single Computer"
    Write-Host "3 - All MS Patches Multiple Computers"
    Write-Host "4 - Specific MS Patch Multiple Computers"
    Write-Host "5 - Exit"
    Write-Host
    
    $Ans = Read-Host "Make Selection"
    
    If ($Ans -eq 1)
    {
        Write-Host
        $Computers = Read-Host "Computer"
        Cls  
        Get-KB
    }
    If ($Ans -eq 2)
    {
        Write-Host
        $Search = Read-Host "Specific KB"
        $Computers = Read-Host "Computer"
        Get-KB
    }
    If ($Ans -eq 3)
    {
        Write-Host
        $Path = Read-Host "Full path to text file with computers"
        $Computers = Get-Content "$Path"
        Cls  
        Get-KB
    }
    If ($Ans -eq 4)
    {
        Write-Host
        $Search = Read-Host "Specific KB"
        $Path = Read-Host "Full path to text file with computers"
        $Computers = Get-Content "$Path"
        Get-KB
    }
}
Until ($Ans -eq 5)