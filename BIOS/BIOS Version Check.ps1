#############Bios function start########### 
function get-Bios 
{ 
param($computername =$env:computername) 
     $L6SBs = "Secure Boot"
     $L6SBv = "Secure Boot"
     $BIOSs = Invoke-Command -ComputerName $Computername {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq "$L6SBs"} | Select-Object -ExpandProperty Name}
     $BIOSv = Invoke-Command -ComputerName $Computername {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq "$L6SBv"} | Select-Object -ExpandProperty CurrentValue}
     $os = Get-WmiObject Win32_bios -ComputerName $computername -ea silentlycontinue 
     If($BIOSv -match "L06 v02.02")
        {
           $servername=$os.PSComputerName  
           $Name= $os.Name
           $SBs=$BIOSs.Name
           $SBv=$BIOSv.CurrentValue
           $results = new-object psobject 
           $results | Add-Member noteproperty PC $servername
           $results | Add-Member noteproperty Setting $SBs
           $results | Add-Member noteproperty Value $SBv 
            
            #Display the results 
            $results | Select-Object PC,BIOSVersion,Manufacturer 
        }else{ 
                 $results = new-object psobject 
                 $results | Add-Member noteproperty PC $servername 
                 $results | Add-Member noteproperty BIOSVersion  $BIOSv
                 $results | Add-Member noteproperty Manufacturer  $Manufacturer  
  
                 #display the results 
                 $results | Select-Object PC,BIOSVersion,Manufacturer
            } 
             
  }

             
####################################Bios function end############################ 
 
#server location  
$servers = Get-Content -Path C:\Temp\2.txt 
$infbios =@() 
foreach($server in $servers)
    { 
        $infbios += get-Bios $server  
    } 
$infbios | export-csv -path c:\Temp\Bios_Versions.csv