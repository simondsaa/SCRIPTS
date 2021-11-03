$PCNames = gc C:\Temp\bios\2.txt 

ForEach($PCNAME in $PCNames){

if (Test-Connection -ComputerName $PCNAME -Count 1 -Quiet) {

$SDC = Invoke-Command -ComputerName $PCNAME {(Get-ItemProperty HKLM:\SOFTWARE\USAF\SDC\ImageRev | Select-Object -ExpandProperty CurrentBuild)}

$NameOfSetting = Invoke-Command -ComputerName $PCNAME {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'Configure Legacy Support and Secure Boot'} | Select-Object -ExpandProperty Name}

$CurrentValue = Invoke-Command -ComputerName $PCNAME {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'Configure Legacy Support and Secure Boot'} | Select-Object -ExpandProperty CurrentValue}

Get-ADComputer $PCNAME -Properties * | Select-Object Name, @{Name = 'SDC Version'; Expression = {$SDC}}, @{Name = 'Name of Setting'; Expression = {$NameOfSetting}}, @{Name = 'Current Value'; Expression = {$CurrentValue}}, Location | Export-Csv -Path 'C:\temp\1online.csv' -NoTypeInformation

} else {

Get-ADComputer $PCNAME -Properties * | Select-Object Name, Location | Export-Csv -Path 'C:\Temp\notonline.csv' -NoTypeInformation -Append

}

}
