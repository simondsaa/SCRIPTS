$Comps = Read-Host "PC List"


Function enableWinRM {
	$result = winrm id -r:$Comps 2>$null

	Write-Host	
	if ($LastExitCode -eq 0) {
		Write-Host "WinRM already enabled on" $Comps "..." -ForegroundColor green
	} else {
		Write-Host "Enabling WinRM on" $Comps "..." -ForegroundColor Yellow
		.\psexec.exe \\$Comps -s C:\Windows\system32\winrm.cmd qc -quiet
		if ($LastExitCode -eq 0) {
			.\psservice.exe \\$Comps restart WinRM
			$result = winrm id -r:$Comps 2>$null

                winrm set winrm/config/client @{TrustedHosts="xlwuw-491s35"}
			
			if ($LastExitCode -eq 0) {Write-Host 'WinRM successfully enabled!' -ForegroundColor green}
			else {exit 1}
		} 
		else {exit 1}
	}
}
enableWinRM
exit 