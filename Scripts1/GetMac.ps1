Function Get-Mac { 

$ComputerName = Get-Content C:\temp\test.txt 

$ErrorActionPreference = 'Stop' 

foreach ($Computer in $ComputerName) { 

Try 

{ 

gwmi -class "Win32_NetworkAdapterConfiguration" -cn $Computer | ? IpEnabled -EQ "True" | 

select DNSHostName, MACAddress | FT -AutoSize 

} 



Catch 

{ 

Write-Warning "System is not reachable : $Computer" 



} 

}#End of Loop 



}#End of the Function 

Get-Mac
