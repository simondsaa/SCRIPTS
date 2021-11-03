$Service = "*masvc*"
$Path = "C:\temp\Mcafee1.txt"
$Servers = Get-Content $Path
ForEach($Server in $Servers){
   Start-Service -cn $Server -ErrorAction SilentlyContinue | Where {$_.Name -like "$Service"} | Select MachineName, Name, Status
   }  
