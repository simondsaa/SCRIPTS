$List = @()

$VMHosts = Get-VMHost

foreach ($VMHost in $VMHosts) {

    $VMHostName = $VMhost.Name

    $esxcli = $VMHost | Get-EsxCli

    $List += $esxcli.software.vib.list() | Select-Object @{N="VMHostName"; E={$VMHostName}}, *

}

$List | Export-Csv -Path C:\Users\1180219788A\Desktop\VIB_List.csv -NoTypeInformation
