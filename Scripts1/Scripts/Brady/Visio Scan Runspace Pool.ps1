$MaxThreads = 200
$SleepTimer = 100
$MaxResultTime = 35

$Start = Get-Date

$Computers = Get-Content "\\XLWUW3-DKPVV1\C$\Users\1392134782A\Desktop\BaseComputers.txt"

$Path = "\\XLWUW3-DKPVV1\C$\Users\1392134782A\Documents\Visio Scan.csv"

If (Test-Path $Path) {Remove-Item $Path}

$ScriptBlock = {
    If (Test-Connection $args[0] -Quiet -Count 1 -BufferSize 16 -Ea 0)
    {
        $Ping = "Online"
        
        Try
        {
            $OSInfo = Get-Wmiobject Win32_OperatingSystem -ComputerName $args[0] -ErrorAction SilentlyContinue
            
            If ($OSInfo.OSArchitecture -eq "64-bit"){$RegPath = "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"}
            ElseIf ($OSInfo.OSArchitecture -eq "32-bit"){$RegPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall"}        
            
            $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$args[0])
            $RegKey = $Reg.OpenSubKey($RegPath)
            $SubKeys = $RegKey.GetSubKeyNames()
            
            ForEach($Key in $SubKeys)
            {
                $ThisKey = $RegPath+"\"+$Key 
                $ThisSubKey = $Reg.OpenSubKey($ThisKey)
                $DisplayNames = $ThisSubKey.GetValue("DisplayName")
                ForEach ($Displayname in $DisplayNames)
                {
                    If ($Displayname -like "Microsoft Visio*")
                    {
                        $Prog = $Displayname
                        $InstallDate = $ThisSubKey.GetValue("InstallDate")
                    }
                }
            }
        }
        
        Catch { $Ping = "No Access" }
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
        Outlook = $Prog
        Install_Date = $InstallDate
        Organization = $Org
        Building = $Bldg
        Room = $Room
        }
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
        Write-Error "Child script appears to be frozen, closing jobs"
        
        ForEach ($Job in $Jobs)
        {
            $Job.Thread.EndInvoke($Job.Handle)
            $Job.Thread.Dispose()
            $Job.Thread = $Null
            $Job.Handle = $Null
        }

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