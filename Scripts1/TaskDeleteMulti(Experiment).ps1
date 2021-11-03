#Experimental TaskScheduler Delete (Multi PCs)

$Comp = Get-Content "C:\temp\tasks.txt"
ForEach ($Computer in $Comp){
$Delete = schtasks.exe /DELETE /TN "JavaT" /S  $Comp /F}