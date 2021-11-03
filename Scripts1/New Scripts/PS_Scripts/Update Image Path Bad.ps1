﻿$LogPath = "C:\Temp"
#Function Fix-ServicePath ([System.IO.DirectoryInfo]$LogPath = "C:\temp") {
<#
	.SYNOPSIS
	    Microsoft Windows Unquoted Service Path Enumeration

	.DESCRIPTION
	    Use Fix-ServicePath to fix vulnerability "Unquoted Service Path Enumeration".
	    
		
	    
	    ------------------------------------------------------------------
	    Author: Vector BCO
		Version: 3.1 
    
	.PARAMETER LogPath	
		You can set different path for log files
		Defaul path is c:\Temp
		Default log file: servicesfix.log
	
	.NOTES


	.EXAMPLE
		 Fix-Servicepath

	.EXAMPLE
		 Fix-ServicePath -LogPath C:\DifferentPath

	.LINK
		https://gallery.technet.microsoft.com/scriptcenter/Windows-Unquoted-Service-190f0341
		https://www.tenable.com/sc-report-templates/microsoft-windows-unquoted-service-path-enumeration
		http://www.commonexploits.com/unquoted-service-paths/
	#>

if (-not (Test-Path $LogPath)){New-Item $LogPath -ItemType directory}

"**************************************************" | Out-File "$LogPath\servicesfix.log" -Append
"Computername: $($Env:COMPUTERNAME)" | Out-File "$LogPath\servicesfix.log" -Append
"Date: $(Get-date -Format "dd.MM.yyyy HH:mm")" | Out-File "$LogPath\servicesfix.log" -Append

# Get all services
Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Services\" | ForEach {
$OriginalPath = (Get-ItemProperty "$($($_).name.replace('HKEY_LOCAL_MACHINE', 'HKLM:'))")
$ImagePath = $OriginalPath.ImagePath

# Get all services with vulnerability
If (($ImagePath -notlike '"*"*') -and ($ImagePath -like '*.exe*'))
{ 
    $NewPath = ($ImagePath -split ".exe ")[0]
    $key = ($ImagePath -split ".exe ")[1]
    $triger = ($ImagePath -split ".exe ")[2]
    $RegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\" + $OriginalPath.PSChildName
    
    Write-Host $ImagePath
    Write-Host $NewPath
    Write-Host $RegPath
    #Write-Host $triger
    # Get all services with vulnerability with key in ImagePath
    If ($NewPath -notlike "*.exe")
    {
        
        " ***** Old Value $ImagePath" | Out-File "$LogPath\servicesfix.log" -Append
        "$($OriginalPath.PSChildName) `"$NewPath.exe`" $key" | Out-File "$LogPath\servicesfix.log" -Append
        #Set-ItemProperty -Path $RegPath -Name "ImagePath" -Value "`"$NewPath.exe`" $key" -Force | Out-Host
    }
        

    # Get all services with vulnerability with out key in ImagePath
    If ($NewPath -like "*.exe")
    {
        
        " ***** Old Value $ImagePath" | Out-File "$LogPath\servicesfix.log" -Append
        "$($OriginalPath.PSChildName) `"$NewPath`"" | Out-File "$LogPath\servicesfix.log" -Append
        #Set-ItemProperty -Path $RegPath -Name "ImagePath" -Value "`"$NewPath`"" -Force | Out-Host
    }
}
}
#}
#Fix-ServicePath -LogPath C:\Temp