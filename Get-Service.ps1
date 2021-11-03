$Service = "*masvc*"
$Path = "C:\temp\ComputersPINGING.txt"
$Servers = Get-Content $Path
ForEach($Server in $Servers){
   Get-Service -cn $Server -ErrorAction SilentlyContinue | Where {$_.Name -like "$Service"} | Select MachineName, Name, Status
   } 