$Computer = Read-Host "Computer Name"
Get-TSSession -ComputerName $Computer -State Active | Stop-TSSession -Force