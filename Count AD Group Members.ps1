$ADGroup = Read-Host "Security Group"
$Number = (Get-ADGroupMember $ADGroup | Measure-Object).Count
Write-Host "$ADGroup has $Number members" 
EXIT