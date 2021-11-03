$Comp = Read-Host "User Name"
    If ($Comp -eq "Pelletier"){$Compname = "XLWUW-491S33"}
    ElseIf ($Comp -eq "Grainger"){$Compname = "XLWUW-491S8K"}
    ElseIf ($Comp -eq "Ballentine"){$Compname = "XLWUW-432LBH"}
    ElseIf ($Comp -eq "Foster"){$Compname = "XLWUW-491S64"}
    ElseIf ($Comp -eq "Mowry"){$Compname = "XLWUW-491S40"}
    ElseIf ($Comp -eq "Lozada"){$Compname = "XLWUW-491S7T"}
    ElseIf ($Comp -eq "Brown"){$Compname = "XLWUW-491S96"}
    ElseIf ($Comp -eq "Barnett"){$Compname = "XLWUW-491S8S"}
    ElseIf ($Comp -eq "Cain"){$Compname = "XLWUW-47168P"}
    ElseIf ($Comp -eq "Simonds"){$Compname = "XLWUW-491S35"}

$User = Get-WmiObject Win32_ComputerSystem -Property Username -Comp $Compname
    If ($User.UserName -eq "AREA52\1383807847N"){$Name = "Pelletier"}
    ElseIf ($User.UserName -eq "AREA52\1253515879N"){$Name = "Grainger"}
    ElseIf ($User.UserName -eq "AREA52\1395576280N"){$Name = "Ballentine"}
    ElseIf ($User.UserName -eq "AREA52\1382931013N"){$Name = "Foster"}
    ElseIf ($User.UserName -eq "AREA52\1383257731N"){$Name = "Mowry"}
    ElseIf ($User.UserName -eq "AREA52\1470230947N"){$Name = "Lozada"}
    ElseIf ($User.UserName -eq "AREA52\1249051671N"){$Name = "Brown"}
    ElseIf ($User.UserName -eq "AREA52\1028801838N"){$Name = "Barnett"}
    ElseIf ($User.UserName -eq "AREA52\1366371229N"){$Name = "Cain"}
    ElseIf ($User.UserName -eq "AREA52\1252862141N"){$Name = "Simonds"}

$Number = Read-Host "Number"
$Phone = "$Number"
$Caller = Read-Host "Caller"
$Subject = Read-Host "Subject"

If (($User.UserName -eq "AREA52\1383807847N") -or 
    ($User.UserName -eq "AREA52\1253515879N") -or 
    ($User.UserName -eq "AREA52\1395576280N") -or 
    ($User.UserName -eq "AREA52\1382931013N") -or 
    ($User.UserName -eq "AREA52\1383257731N") -or
    ($User.UserName -eq "AREA52\1470230947N") -or
    ($User.UserName -eq "AREA52\1249051671N") -or
    ($User.UserName -eq "AREA52\1028801838N") -or
    ($User.UserName -eq "AREA52\1252862141N") -or
    ($User.UserName -eq "AREA52\1366371229N"))
    {$Message = "From: Simonds, Aaron C TSgt 101 ACOMS/SCOO

You had a missed call from $Caller @ $Phone.

Subject: $Subject"
    Msg Console /Server:$Compname $Message
    Write-Host
    Write-Host "User Messaged: $Name"}
Else {Write-Host "The specified user is not logged on. Current user: $User" $User.UserName}