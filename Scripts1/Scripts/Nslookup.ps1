$IPs = Get-Content 'C:\Users\1252862141.adm\Desktop\Scripts1\Pop.txt'

ForEach ($IP in $IPs)
{
    Try { $Name = [System.Net.DNS]::GetHostByAddress($IP).HostName.Split(".")[0] }
    Catch { $Name = "No record" }
    Write-Host "$IP : $Name"
}