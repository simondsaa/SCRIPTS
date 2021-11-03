#this finally accesses HP BIOS successfully. i need it to print the PC name it's representing.
#function Get-HPBIOSSettings
#{ 
$Array = @()
$N = 'Name'
$CV = 'CurrentValue'
$PV = 'PossibleValue'
$S = 'Secure'
$L = 'Legacy*'
$SB = 'SecureBoot'
$Path = "C:\temp\2.txt"
$Computers = gc $Path
foreach ($Computer in $Computers) 
    {
    $BIOS = Get-WmiObject -computername $Computer -Namespace root/hp/instrumentedBIOS -Class HP_BIOSEnumeration | select-object Name, CurrentValue 
                
                $obj = New-Object PSObject
                $obj | Add-Member -Force -MemberType NoteProperty -Name "Computer" -Value $Computer
            
    $BIOS | ForEach{
                 if($BIOS.CurrentSetting -ne ""){ 
        $Setting = $BIOS.CurrentSetting -split ',' 
        $obj | Add-Member -MemberType NoteProperty -Name "yep" -Value $Setting -Force       
    }
  }
$obj
}

