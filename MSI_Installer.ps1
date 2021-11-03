#Variables

$computername = Get-Content 'C:\Temp\SwiftPCs.txt'

$sourcefile = "C:\Temp\Software\Swift 4_0_2\Swift-4.0.2.msi"

#This section will install the software 

foreach ($computer in $computername) 

{

$destinationFolder = "\\$computer\C$\Temp"

#This section will copy the $sourcefile to the $destinationfolder. If the Folder does not exist it will create it.

if (!(Test-Path -path $destinationFolder))

{

New-Item $destinationFolder -Type Directory

}

Copy-Item -Path $sourcefile -Destination $destinationFolder

Invoke-Command -ComputerName $computer -ScriptBlock { & cmd /c "msiexec.exe /i c:\Temp\Swift-4.0.2.msi" /qn ADVANCED_OPTIONS=1 CHANNEL=100}

}