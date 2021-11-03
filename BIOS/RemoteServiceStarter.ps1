<#
    .Synopsis 
        Restart a service on list of remote computers.
        
    .Description
        This script helps in restarting a service remotely on list of remote computers.
 
    .Parameter ComputerName    
        Computer name(s) for which you want to get the disk space details.
        
    .Example
        Restart-Service.ps1 -ComputerName Comp1, Comp2 -ServiceName dnscache
		
		Restart DNSCache service on Comp1 and Comp2 computers and report the status
#>

[cmdletbinding()]
param(
	[parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
	[string[]]$ComputerName = $env:computername,
	
	[parameter(Mandatory=$true)]
	[string]$ServiceName,
	
	[string]$OutputDir = "C:\Temp\RemoteRegistrySTATUS"
)

begin {
}
process{
    $ComputerName = gc "C:\Temp\G2.txt"
	$SuccessComputers  = Join-Path $OutputDir "SuccessComputers.csv"
	$FailedComputers   = join-path $OutputDir "FailedComputers.csv"
	$OutputArray = @()
	foreach($Computer in $ComputerName) {
		$OutputObj	= New-Object -TypeName PSobject 
		$OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.TOUpper()
		Write-Verbose "Working on $Computer"
		$Status = "Failed"
		$IsOnline=$false
		if(Test-Connection -ComputerName $Computer -Count 1 -ea 0) {
			$IsOnline = $true
			try {
				$ServiceObj = Get-Service -Name $ServiceName -ComputerName $Computer -ErrorAction Stop
				Restart-Service -InputObj $ServiceObj -erroraction stop
				$Status="Running"
				
			} catch {
				Write-Verbose "Failed to restart $Service on $Computer. Error: $_"
				$Status="Failed"
			}
			
			
		}
		else {
			Write-Verbose "$Computer is not reachable"
			$IsOnline = $false
			
		}
		$OutputObj | Add-Member -MemberType NoteProperty -Name Status -Value $Status
		$OutputObj | Add-Member -MemberType NoteProperty -Name IsOnline -Value $IsOnline
		$OutputObj
		$OutputArray += $OutputObj
	}

	$OutputArray | ? {$_.Status -eq "Failed" -or $_.IsOnline -eq $false} | Out-File -FilePath $FailedComputers
	$OutputArray | ? {$_.Status -eq "Running"} | Out-File -FilePath $SuccessComputers
}
end {
}
