Write-Host
$Time = Read-Host "How long (seconds)"
Write-Host
$Comp = Read-Host "User Name: "
If ($Comp -eq "Tobler"){$Compname = "52XLWUW3-DLKVV1"}
ElseIf ($Comp -eq "Brady"){$Compname = "52XLWUW3-DKPVV1"}
ElseIf ($Comp -eq "Lonnie"){$Compname = "52XLWUW3-DJPVV1"}
ElseIf ($Comp -eq "Trey"){$Compname = "52XLWUW3-DDMVV1"}
ElseIf ($Comp -eq "Le"){$Compname = "52XLWUW3-431YVS"}
ElseIf ($Comp -eq "Roberson"){$Compname = "52XLWUW3-DJKVV1"}
ElseIf ($Comp -eq "Wilson"){$Compname = "52XLWUW3-DJHVV1"}
ElseIf ($Comp -eq "Carlos"){$Compname = "52XLWUW3-DHMVV1"}
ElseIf ($Comp -eq "Jones"){$Compname = "52XLWUW3-DLJVV1"}
$User = Get-WmiObject Win32_ComputerSystem -Property Username -Comp $Compname
If ($User.UserName -eq "AREA52\1129967354A"){$Name = "Carlos"}
ElseIf ($User.UserName -eq "AREA52\1407477894A"){$Name = "Tobler"}
ElseIf ($User.UserName -eq "AREA52\1274873341C"){$Name = "Lonnie"}
ElseIf ($User.UserName -eq "AREA52\1258114554C"){$Name = "Trey"}
ElseIf ($User.UserName -eq "AREA52\1392134782A"){$Name = "Brady"}
ElseIf ($User.UserName -eq "AREA52\1381589257A"){$Name = "Le"}
ElseIf ($User.UserName -eq "AREA52\1180219788A"){$Name = "Roberson"}
ElseIf ($User.UserName -eq "AREA52\1455880293A"){$Name = "McNeal"}
ElseIf ($User.UserName -eq "AREA52\1186019462C"){$Name = "Jones"}
If (($User.UserName -eq "AREA52\1407477894A") -or 
    ($User.UserName -eq "AREA52\1274873341C") -or 
    ($User.UserName -eq "AREA52\1258114554C") -or 
    ($User.UserName -eq "AREA52\1392134782A") -or 
    ($User.UserName -eq "AREA52\1381589257A") -or
    ($User.UserName -eq "AREA52\1180219788A") -or
    ($User.UserName -eq "AREA52\1129967354A") -or
    ($User.UserName -eq "AREA52\1455880293A") -or
    ($User.UserName -eq "AREA52\1186019462C"))
    {Write-Host
    $Message = Read-Host "Message"
    Shutdown /r /f /m \\$Compname /t $Time /c "$Message"
    Write-Host
    Write-Host "User Messaged: $Name"}
Else {Write-Host "The specified user is not logged on. Current user: $User" $User.UserName}