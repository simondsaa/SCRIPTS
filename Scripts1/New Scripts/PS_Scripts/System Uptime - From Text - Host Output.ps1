$Computers = Get-Content "C:\Users\timothy.brady\Desktop\Servers.txt"
ForEach ($computer in $Computers)
{
    If (Test-Connection $computer -quiet -BufferSize 16 -Ea 0 -count 2)
    {
    $lastboottime = (Get-WmiObject Win32_OperatingSystem -cn $computer -ErrorAction SilentlyContinue).LastBootUpTime
    $sysuptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($lastboottime) 
        If ($sysuptime.Days -gt 7 -and $sysuptime.Days -lt 30){$color = "Yellow"}
        ElseIf ($sysuptime.Days -gt 30){$color = "Red"}
        Else {$color = "White"}
        Write-Host -ForegroundColor $color "$computer has been up for: " $sysuptime.days "days" $sysuptime.hours "hours" $sysuptime.minutes "minutes" $sysuptime.seconds "seconds"
    }
    Else {Write-Host -ForegroundColor Black "$computer is not reachable"}
}