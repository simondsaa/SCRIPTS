$Path = "C:\Temp\G2.txt"
$Computers = gc $Path
$property =  "" | select Name, Currentvalue, possiblevalues
foreach ($Computer in $Computers) 
    {
       Get-WmiObject -computername $Computer -Namespace root/hp/instrumentedBIOS -Class HP_BIOSEnumeration | select-object Name, CurrentValue, PossibleValues | Where-Object {($_.Name -like "Secure*") -or ($_.Name -like "Legacy*") -or ($_.Name -like "SecureBoot")}
    }
    #Possible Lenovo
    #Get-WmiObject -class Lenovo_BiosSetting -namespace root\wmi -ComputerName $Computer | select-object Name, CurrentValue, PossibleValues | Where-Object {($_.Name -like "Secure*") -or ($_.Name -like "Legacy*") -or ($_.Name -like "SecureBoot")}
