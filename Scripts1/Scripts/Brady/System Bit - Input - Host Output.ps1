$Computer = Read-Host "Computer Name"
If (Test-Connection $Computer -quiet -count 1)
{    
    $Info = Get-WmiObject -ComputerName $Computer Win32_ComputerSystem
    Write-Host
    Write-Host $Info.Name":"$Info.SystemType
}
Else
{
    Write-Host
    Write-Host "$Computer : unavailable"
}