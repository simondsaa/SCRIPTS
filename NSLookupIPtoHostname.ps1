$IPs = Get-Content 'C:\temp\PINGTEST.txt'

ForEach ($IP in $IPs)
{
    Try { $Name = [System.Net.DNS]::GetHostByAddress($IP).HostName.Split(".")[0] }
    Catch { $Name = "No record" }
    Write-Host "$IP : $Name"
}