$MaxConcurrentJobs = 50
               ##Insert target list######################
$strFilename = "C:\Users\timothy.brady\Desktop\Comps.txt"
if(test-path $strFilename) {
    $list = (get-content $strFilename)
    "Approx {0} Systems for scan" -f ($list | Measure-Object).count
foreach($_ in $list) {
    if ($_.length -gt 0)
    {   "Scanning {0}" -f $_                                   #Insert KB below#
        start-job -scriptblock { invoke-expression "get-hotfix -id 'KB2737019' -computername '$args'"} -name("Scan -" + $_) -argumentlist $_ | out-null}
    while (((get-job | where-object { $_.Name -like "Scan*" -and $_.State -eq "Running"}) | measure).Count -gt $MaxConcurrentJobs)
    {   "{0} Concurrent jobs running, sleeping 5 seconds" -f $MaxConcurrentJobs
        Start-Sleep -seconds 5}
                }
    While (((get-job | where-object { $_.Name -like "Scan*" -and $_.State -eq "Running" }) | measure).Count -gt 0)
    {   $jobcount =  ((get-job | where-object { $_.Name -like "Scan*" -and $_.State -eq "Running" }) | measure).Count
        write-host "Waiting for $jobcount Jobs to Complete"
        Start-Sleep -seconds 5
        $Counter++
        if ($Counter -gt 40)
        {   write-host "Exiting loop $jobCount Jobs did not complete"
            get-job | where-object { $_.Name -like "Ping*" -and $_.State -eq "Running" } | select Name
            break
        }
     }

$PingResults = @( )

get-job | where { $_.Name -like "Scan*" -and $_.State -eq "Completed" } | % { $PingResults += Receive-Job $_ ; Remove-Job $_ }

Write-Host
Write-Host "Scan Complete!"
"Total time for scan: {0} Minutes" -f [math]::round(($stoptimer - $starttimer).TotalMinutes , 2)
"{0} Systems Offline, and did not connect" -f (($PingResults | where-object { $_.HotFixID -eq $Null}) | measure-object ).Count
"{0} Good Systems" -f (($PingResults | where-object {$_.HotFixID -like "KB2737019"}) | measure-object ).Count
}                                                                 #Insert KB Above#  
else
{"Invalid FilePath!"}
$a = New-Object -comobject Excel.Application
$a.visible = $True 
$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)
$c.Cells.Item(1,1) = "Computer Name:"
$c.Cells.Item(1,2) = "KB2737019:"
$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True
$d.EntireColumn.AutoFilter() | out-null
$d.EntireColumn.AutoFit() | out-null
$intRow = 2
foreach($strComputer in $list)
{$c.Cells.Item($intRow,1) = $strComputer
$Out = $PingResults.GetType().CSName
if($PingResults -match $strComputer)
{$c.Cells.Item($intRow,2).Interior.ColorIndex = 4
$c.Cells.Item($intRow,2) = "Good"}
else
{$c.Cells.Item($intRow,2).Interior.ColorIndex = 3
$c.Cells.Item($intRow,2) = "Bad"}
$intRow = $intRow + 1}
$date = Get-Date
$filename = "\WIN7KB2617657_{0}{1:d2}{2:d2}-{3:d2}{4:d2}" -f $date.year,$date.month,$date.day,$date.hour,$date.minute
$filename = $filename + ".xlsx"
$location = get-location
$filename = $location.path + $filename
Write-host "Saving file in:"
write-host $filename
$c.SaveAs("$filename")