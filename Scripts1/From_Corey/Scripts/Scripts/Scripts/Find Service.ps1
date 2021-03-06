#Service to search for
$Service = "SNMP"

#Path to your list of servers
$Servers = Get-Content "C:\Users\timothy.brady\Desktop\Servers.txt"

#Command starts
ForEach($Server in $Servers){
   Get-Service -cn $Server -ErrorAction SilentlyContinue | Where {$_.Name -like "$Service"} | Select MachineName, Name, Status
   } 