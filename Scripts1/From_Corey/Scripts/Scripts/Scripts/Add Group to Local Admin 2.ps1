$Computers = Get-Content "C:\USers\timothy.brady\Desktop\Comps.txt" |
ForEach {If (Net localgroup administrators) {Write-Host $_.Group}
Else {Write-Host "Group not there"}}