
[cmdletbinding()]
param(
	[parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
	[string[]]$ComputerName = $env:computername,
	
	[string]$OutputDir = "C:\Temp\BIOS TEST SCRIPT"
)

begin {
}
process{
    $ComputerName = gc "C:\Temp\2.txt"
	$OutputArray = @()
	foreach($Computer in $ComputerName) {
		if(Test-Connection -ComputerName $Computer -Count 1 -ea 0) {
			try {
                $Script:Get_BIOS_Settings = Get-WmiObject -computername $Computer -Namespace root/hp/instrumentedBIOS -Class HP_BIOSEnumeration |  % { New-Object psobject -Property @{
                Setting = $_."Name" 
                Value = $_."currentvalue"
                Available_Values = $_."possiblevalues"
                }} | select-object Name, Value, PossibleValues 
                $OutputObj | Add-Member -Force -MemberType NoteProperty -Name ComputerName -Value $Computer
				
			} catch {
				Write-Verbose "Failed to pull BIOS settings on $Computer. Error: $_"
			}
					
		}
		else {
			Write-Host "$Computer is not reachable"
			
		}
        
    ForEach($obj in $Script:Get_BIOS_Settings){
        $OutputObj | Add-Member -Force -MemberType NoteProperty -Name Setting -Value $obj
        $OutputObj | Add-Member -Force -MemberType NoteProperty -Name Value -Value $obj
        $OutputObj | Add-Member -Force -MemberType NoteProperty -Name Possible_Values -Value $obj
        }
		$OutputArray += $OutputObj
}
    $OutputArray | Select ComputerName, Setting, Value, Possible_Values
}
end {
}
