#############Bios function start########### 
function get-Bios { 
 param( 
 $computername =$env:computername 
 ) 
 
 $os = Get-WmiObject Win32_bios -ComputerName $computername -ea silentlycontinue 
 if($os){ 
 
   $SerialNumber =$os.SerialNumber 
   $servername=$os.PSComputerName  
   $Name= $os.Name 
   $SMBIOSBIOSVersion=$os.SMBIOSBIOSVersion 
   $Manufacturer=$os.Manufacturer 
 
 
 
 
 $results =new-object psobject 
 
 $results |Add-Member noteproperty SerialNumber  $SerialNumber 
 $results |Add-Member noteproperty ComputerName  $servername 
 $results |Add-Member noteproperty Name  $Name 
 $results |Add-Member noteproperty SMBIOSBIOSVersion  $SMBIOSBIOSVersion 
 $results |Add-Member noteproperty Manufacture   $Manufacture 
  
 
 
 #Display the results 
 
 $results | Select-Object computername,SMBIOSBIOSVersion,Name,Manufacture ,SerialNumber 
 
 } 
 
 
 else 
 
 { 
 
  
 $results =new-object psobject 
 
 $results |Add-Member noteproperty SerialNumber "Na" 
 $results |Add-Member noteproperty ComputerName $servername 
  $results |Add-Member noteproperty Name  $Name 
 $results |Add-Member noteproperty SMBIOSBIOSVersion  $SMBIOSBIOSVersion 
 $results |Add-Member noteproperty Manufacture   $Manufacture 
 
 
  
 #display the results 
 
 
 $results | Select-Object computername,SMBIOSBIOSVersion,Name,Manufacture ,SerialNumber 
 
 
 
 
 } 
 
 
 
 } 
 
 
 
 ####################################Bios function end############################ 
 
 #server location  
 
 $servers = Get-Content -Path C:\Temp\2.txt 
 
 $infbios =@() 
 
 
 foreach($server in $servers){ 
 
 $infbios += get-Bios $server  
 } 
 
 $infbios | export-csv -path c:\Temp\Bios_Versions.csv