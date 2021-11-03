$MaxThreads = 200
$SleepTimer = 100
$MaxResultTime = 35

$Start = Get-Date

$Computers = Get-Content "C:\Users\1180219788A\Desktop\Computers.txt"

$Path = "C:\Users\1180219788A\Desktop\64Bit.csv"

If (Test-Path $Path) {Remove-Item $Path}

$ScriptBlock = {
    If (Test-Connection $args[0] -Quiet -Count 1 -BufferSize 16 -Ea 0)
    {
        $Ping = "Online"
        Try
        {
            $Bit = (Get-WmiObject Win32_OperatingSystem -cn $args[0] -ErrorAction SilentlyContinue).OSArchitecture
            $RAM = [Math]::Round((Get-WmiObject Win32_ComputerSystem -cn $args[0] -ErrorAction SilentlyContinue).TotalPhysicalMemory/1048576, 0)
            $Disk = Get-WmiObject Win32_LogicalDisk -cn $args[0] -ErrorAction SilentlyContinue | Where {$_.DeviceID -like "C:"}
            $FreeSpace = [Math]::Round($Disk.FreeSpace/1073741824, 0)
        }
        Catch { $Ping = "No Access" }
        
        $Profiles = (Get-ChildItem "\\$($args[0])\C$\Users").Count
    }
    Else
    {
        $Ping = "Offline"
    }

    $Domain = "OU=Tyndall AFB Computers,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
    $objDomain = [adsi]("LDAP://" + $domain)
    $Search = New-Object System.DirectoryServices.DirectorySearcher
    $Search.SearchRoot = $objDomain
    $Search.Filter = "(&(objectClass=computer)(samAccountName=*$($args[0])*))"
    $Search.SearchScope = "Subtree"
    $Results = $Search.FindAll()
    ForEach($Item in $Results)
    {
        $objComputer = $Item.GetDirectoryEntry()
        $Org = (($objComputer.o) | Out-String).Trim()
        $Bldg = ($objComputer.location).Split(";")[0]
        $Room = ($objComputer.location).Split(";")[1].TrimStart(" ")
    }

    [PSCustomObject]@{
        System = $args[0]
        Ping = $Ping
        SystemBit = $Bit
        RAM_MB = $RAM
        FreeDiskSpace_GB = $FreeSpace
        Organization = $Org
        Building = $Bldg
        Room = $Room
        Profiles = $Profiles
        }
    
    #$Results
}

$ISS = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $MaxThreads, $ISS, $Host)
$RunspacePool.Open()
        
$Jobs = @()

Write-Progress -Activity "Preloading threads" -Status "Starting Job $($jobs.count)"
ForEach ($Computer in $Computers)
{
    $PowershellThread = [PowerShell]::Create().AddScript($ScriptBlock)
    $PowershellThread.RunspacePool = $RunspacePool
    $PowershellThread.AddArgument($Computer.ToString()) | Out-Null
    $Handle = $PowershellThread.BeginInvoke()
    $Job = "" | Select-Object Handle, Thread, object
    $Job.Handle = $Handle
    $Job.Thread = $PowershellThread
    $Job.Object = $Computer.ToString()
    $Jobs += $Job
}

$ResultTimer = Get-Date
    
While (@($Jobs | Where-Object {$_.Handle -ne $Null}).count -gt 0)
{
    $Remaining = "$($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).object)"
    
    If ($Remaining.Length -gt 60)
    {
        $Remaining = $Remaining.Substring(0,60) + "..."
    }
    
    Write-Progress `
    -Activity "Waiting for Jobs - $($MaxThreads - $($RunspacePool.GetAvailableRunspaces())) of $MaxThreads threads running" `
    -PercentComplete (($Jobs.count - $($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).count)) / $Jobs.Count * 100) `
    -Status "$(@($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})).count) jobs remaining" 
 
    ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $True}))
    {
        ($Job.Thread.EndInvoke($Job.Handle)) | Export-Csv $Path -Append -NoTypeInformation -Force
        $Job.Thread.Dispose()
        $Job.Thread = $Null
        $Job.Handle = $Null
        $ResultTimer = Get-Date
    }
        
    If (($(Get-Date) - $ResultTimer).totalseconds -gt $MaxResultTime)
    {
            Write-Error "Child script appears to be frozen, try increasing MaxResultTime"
            Break 
    }
    
    Start-Sleep -Milliseconds $SleepTimer
} 

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
$ScriptTime = "Elapsed Time: $TimeM min $TimeS sec"
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan
Write-Host
Write-Host "Closing runspace pools..."

$RunspacePool.Close()
$RunspacePool.Dispose()