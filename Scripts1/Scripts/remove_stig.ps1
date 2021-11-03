$vibs = @("dod-esxi6-stig-rd")
Get-vmhost | foreach {
Write-host "Working on Host: $_"
$esxcli = get-esxcli -vmhost $_
foreach ($vib in ($vibs)) {
write-host "      searching for vib $vib" -ForegroundColor Cyan
    if ($esxcli.software.vib.get.invoke() | where {$_.name -eq "$vib"} -erroraction silentlycontinue )  {
        write-host "      found vib $vib. Deleting" -ForegroundColor Green
        $esxcli.software.vib.remove.invoke($null, $true, $false, $true, "$vib") 
    } else {
        write-host "      vib $vib not found. continuing..." -ForegroundColor Yellow
    }
}
}