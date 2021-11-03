$ADGroup = Read-Host "Security Group"
$Number = (Get-ADGroupMember $ADGroup | Measure-Object).Count
Write-Host "$ADGroup has $Number members" 
Write-Host "Press any key to Exit..."
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp") > $null
EXIT