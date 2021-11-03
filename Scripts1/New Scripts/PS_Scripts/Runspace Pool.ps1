$Start = Get-Date

$Computers = Get-Content "\\XLWUW3-DKPVV1\C$\Users\1392134782A\Desktop\BaseComputers.txt"

$Counter = $null

$scriptblock = 
{
    Param (
    [string]$Computer
    )

    If (Test-Connection $Computer -Quiet -Count 1 -BufferSize 16 -Ea 0)
    {
        $Ping = "Online"
        Try
        {
            $Bit = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).OSArchitecture
            $RAM = [Math]::Round((Get-WmiObject Win32_ComputerSystem -cn $Computer -ErrorAction SilentlyContinue).TotalPhysicalMemory/1048576, 0)
            $Disk = Get-WmiObject Win32_LogicalDisk -cn $Computer -ErrorAction SilentlyContinue | Where {$_.DeviceID -like "C:"}
            $FreeSpace = [Math]::Round($Disk.FreeSpace/1073741824, 0)
        }
        Catch { }
        
        $Profiles = (Get-ChildItem "\\$Computer\C$\Users").Count
    }
    Else
    {
        $Ping = "Offline"
    }

    [PSCustomObject]@{
    System = $Computer
    Ping = $Ping
    SystemBit = $Bit
    RAM_MB = $RAM
    FreeDiskSpace_GB = $FreeSpace
    Profiles = $Profiles
    }
}

$Path = "\\XLWUW3-DKPVV1\C$\Users\1392134782A\Documents\Runspace Resutls.csv"

$RunspacePool = [RunspaceFactory]::CreateRunspacePool(100,100)
$RunspacePool.Open()
$Jobs = 
    ForEach ($Computer in $Computers)
{
     $Job = [PowerShell]::Create().
            AddScript($ScriptBlock).
            AddArgument($Computer)
     $Job.RunspacePool = $RunspacePool

     [PSCustomObject]@{
      Pipe = $Job
      Result = $Job.BeginInvoke()
     }
}

Write-Host 'Working...' -NoNewline

Do {
   $Counter++
   Write-Host $Counter
   Start-Sleep -Seconds 1
} While ( $Jobs.Result.IsCompleted -contains $false)

Write-Host ' Done! Writing output file.'
Write-host "Output file is $Path"

$(ForEach ($Job in $Jobs)
{ $Job.Pipe.EndInvoke($Job.Result) }) |
 Export-Csv $Path -NoTypeInformation

$RunspacePool.Close()
$RunspacePool.Dispose()

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan