$MaxThreads = 20
$SleepTimer = 200
$MaxResultTime = 3000
[HashTable]$AddParam = @{}
[Array]$AddSwitch = @()

$Computers = Get-Content "\\XLWUW3-DKPVV1\C$\Users\1392134782A\Desktop\Comps.txt"

$Path = "\\XLWUW3-DKPVV1\C$\Users\1392134782A\Documents\64Bit.csv"

$ScriptBlock = {
    
    Param (
    [string]$Computer
    )
    
    If (Test-Connection $Computer -Quiet -Count 1 -BufferSize 16 -Ea 0)
    {
        $Ping = "Online"
        Try
        {
            $Bit = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).OSArchitecture
            $RAM = [Math]::Round((Get-WmiObject Win32_ComputerSystem -cn $args[0] -ErrorAction SilentlyContinue).TotalPhysicalMemory/1048576, 0)
            $Disk = Get-WmiObject Win32_LogicalDisk -cn $args[0] -ErrorAction SilentlyContinue | Where {$_.DeviceID -like "C:"}
            $FreeSpace = [Math]::Round($Disk.FreeSpace/1073741824, 0)
        }
        Catch { }
        
        $Profiles = (Get-ChildItem "\\$($args[0])\C$\Users").Count
    }
    Else
    {
        $Ping = "Offline"
    }

    #$AD = Get-ADComputer -Identity $args[0] -Properties location, o
    #$Org = $AD.o
    #$Bldg = ($AD.location).Split(";")[0]
    #$Room = ($AD.location).Split(";")[1]

    [PSCustomObject]@{
        System = $Computer
        Ping = $Ping
        SystemBit = $Bit
        RAM_MB = $RAM
        FreeDiskSpace_GB = $FreeSpace
        #Organization = $Org
        #Building = $Bldg
        #Room = $Room
        Profiles = $Profiles
        }
    
    #$Results
}

$ISS = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $MaxThreads, $ISS, $Host)
$RunspacePool.Open()
        
Write-Progress -Activity "Preloading threads" -Status "Starting Job $($jobs.count)"
$Jobs =
ForEach ($Computer in $Computers)
{
    $Job = [PowerShell]::Create().AddScript($ScriptBlock)
    $Job.AddArgument($Computer)
    $Job.RunspacePool = $RunspacePool
    $Job.Handle = $Job.BeginInvoke()

    [PSCustomObject]@{
        Pipe = $Job
        Result = $Job.BeginInvoke()
        }
    
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
    -Status "$(@($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})).count) remaining - $remaining" 
 
    $(ForEach ($Job in $Jobs)
    {
        $Job.Pipe.EndInvoke($Job.Result)
    }) | Export-Csv $Path -NoTypeInformation
    
    ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $True}))
    {
        $Job.Thread.EndInvoke($Job.Handle)
        $Job.Thread.Dispose()
        $Job.Thread = $Null
        $Job.Handle = $Null
        $ResultTimer = Get-Date
    }
        
    If (($(Get-Date) - $ResultTimer).totalseconds -gt $MaxResultTime)
    {
            Write-Error "Child script appears to be frozen, try increasing MaxResultTime"
            Exit
    }
    
    Start-Sleep -Milliseconds $SleepTimer
} 

$RunspacePool.Close() | Out-Null
$RunspacePool.Dispose() | Out-Null