#this finally accesses HP BIOS successfully. i need it to print the PC name it's representing.
function Get-HPBIOSSettings
{ 

[cmdletbinding()]
param(
	[parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
	[string[]]$ComputerName = $env:computername,
	
	[string]$OutputDir = "C:\Temp\BIOS TEST SCRIPT"
)

begin {
}
process{
    $S = 'Secure'
    $L = 'Legacy*'
    $SB = 'SecureBoot'
    $ComputerName = gc "C:\Temp\2.txt"
	$SuccessComputers  = Join-Path $OutputDir "BIOS PULL SUCCESS.csv"
	$FailedComputers   = join-path $OutputDir "BIOS PULL FAILED.csv"
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
                Get-WmiObject -computername $Computer -Namespace root/hp/instrumentedBIOS -Class HP_BIOSEnumeration | select-object Name, CurrentValue, PossibleValues | Where-Object {($_.Name -like "$S") -or ($_.Name -like "$L") -or ($_.Name -like "$SB")}
				$Status="Running"
				
			} catch {
				Write-Verbose "Failed to pull BIOS settings on $Computer. Error: $_"
				$Status="Failed"
			}
					
		}
		else {
			Write-Verbose "$Computer is not reachable"
			$IsOnline = $false
			
		}
        $OutputObj | Add-Member -MemberType NoteProperty -Name Name -Value $S
        $OutputObj | Add-Member -MemberType NoteProperty -Name Current_Value -Value $L
        $OutputObj | Add-Member -MemberType NoteProperty -Name Possible_Values -Value $SB
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
}