$Path = "C:\Temp\G2.txt"
$Computers = gc $Path
$Comp = gc $Path
ForEach($Comp in $Computers){
Invoke-Command -ComputerName $Comp -ArgumentList $servicename -ScriptBlock $SBlock 
}
$SBlock = {
    Param($servicename)
    $ServiceName = "RemoteRegistry"
    $Service = Get-Service -name $servicename
    if ($Service.Status -eq "Running"){
        Write-Host "The RemoteRegistry service is started on $Comp"
    }
    Else{
        Write-Host "The RemoteRegistry service is stopped on $Comp, starting up service"
        start-sleep -seconds 5
        Start-Service -name $servicename
        Write-Host "The RemoteRegistry service is starting $Comp"
        start-sleep -seconds 10
        $Service.Refresh()
        if ($server1.status -eq "Running"){
            Write-Host "The RemoteRegistry service is now started on $Comp"
        }
        else {
           Write-Host "The RemoteRegistry service failed to start on $Comp please check services and try again."
        }
    }
}
