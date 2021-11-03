            $directory = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Current\Computer_stats\Windows\*.*"
            $csvFiles = get-childitem $directory -filter *.csv

            $results = @();


            Foreach ($csv in $csvFiles) {
                $results += import-csv $csv
                $i++
                Write-Host "." -ForegroundColor Cyan
                Write-Progress -activity “Combining Information” -status “Status: $i/$($csvfiles.count)” -PercentComplete (($i / $csvFiles.count)*100)
                 }

            $i = $null

# Path to your computer names you'd like to match to Stats
$compList = get-content "\\xlwu-fs-05pv\Tyndall_PUBLIC\offlines.txt"


$match = $results | where {$complist -contains $_.computername} | select Date, ComputerName, First.Last, EDIPI, Organization,OS_Version,OS_Architecture,server,netlogon,remoting,ipaddress,mac,defaultIPgateway,dhcp,manufactrer,serialnumber,location,dayofyear

"Match count : $($match.count)"
$match | ogv