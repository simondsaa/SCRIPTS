Do
{
    Cls
    Write-Host " "
    Write-Host "Which Stats to display?"
    Write-Host "1 - Users"
    Write-Host "2 - Computers"
    Write-Host "3 - Exit"
    Write-Host " "

    $Ans = Read-Host "Make Selection"
    
    If ($Ans -eq 1)
    {

cls
# User Selection
# Create blank line
$nl = [Environment]::newline

Write-Host "Computer Stat Joiner" -ForegroundColor Red
Write-Host "Combines information for all current logs related to computer accounts,"
write-Host "for a hisotry, please use the archive/Concurrent files."
write-host "NOTE :: This will take at least 40 seconds for a full data pull" -ForegroundColor DarkYellow





$directory = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Current\User_Stats\*.*"
$csvFiles = get-childitem $directory -filter *.csv
 
#Progress counter
$i=0

$results = @();
Foreach ($csv in $csvFiles) {
    $results += import-csv $csv
    $i++
    Write-Host "." -ForegroundColor Cyan
    Write-Progress -activity “Combining Information” -status “Status: $i/$($csvfiles.count)” -PercentComplete (($i / $csvFiles.count)*100)
        }
        

Do
{
    Cls
    $nl
    $nl
    $nl
    $nl
    Write-Host " "
    Write-Host "Please select how you would like the information Displayed."
    Write-Host "1 - PowerShell OGV (No Export) (Can Manipulate)"
    Write-Host "2 - Excel (Export) (Can Manipulate)"
    Write-Host "3 - Exit"
    Write-Host " "

    $Ans = Read-Host "Make Selection"
    
    If ($Ans -eq 1)
    {
        $results | ogv
    }
    If ($Ans -eq 2)
    {
        $results | Export-Csv "C:\Windows\Temp\User_Stats.csv"
        Invoke-Item "C:\Windows\Temp\User_Stats.csv"
}
}
Until ($Ans -eq 3)

}

    
    If ($Ans -eq 2)
    {

# Computer Selection
# Create blank line
$nl = [Environment]::newline

Write-Host "Computer Stat Joiner" -ForegroundColor Red
Write-Host "Combines information for all current logs related to computer accounts,"
write-Host "for a hisotry, please use the archive/Concurrent files."
write-host "NOTE :: This will take at least 40 seconds for a full data pull" -ForegroundColor DarkYellow





$directory = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Current\Computer_Stats\Windows\*.*"
$csvFiles = get-childitem $directory -filter *.csv
 
#Progress counter
$i=0

$results = @();
Foreach ($csv in $csvFiles) {
    $results += import-csv $csv
    $i++
    Write-Host "." -ForegroundColor Cyan
    Write-Progress -activity “Combining Information” -status “Status: $i/$($csvfiles.count)” -PercentComplete (($i / $csvFiles.count)*100)
        }
        

Do
{
    Cls
    $nl
    $nl
    $nl
    $nl
    Write-Host " "
    Write-Host " Active Remoting Enabled: $(($results | where {$_.Remoting -like "Active"}).count)" -ForegroundColor Green
    Write-Host " Active Remoting Disabled: $(($results | where {$_.Remoting -like "Inactive"}).count)" -ForegroundColor Yellow
    Write-Host " Active Netlogon Enabled: $(($results | where {$_.Netlogon -like "Active"}).count)" -ForegroundColor Green
    Write-Host " Active Netlogon Disabled: $(($results | where {$_.Netlogon -like "Inactive"}).count)" -ForegroundColor Yellow
    Write-Host "Please select how you would like the information Displayed."
    Write-Host "1 - PowerShell OGV (No Export) (Can Manipulate)"
    Write-Host "2 - Excel (Export) (Can Manipulate)"
    Write-Host "3 - Exit"
    Write-Host " "

    $Ans = Read-Host "Make Selection"
    
    If ($Ans -eq 1)
    {
        $results | ogv
    }
    If ($Ans -eq 2)
    {
        $results | Export-Csv "C:\Windows\Temp\Comp_Stats.csv"
        Invoke-Item "C:\Windows\Temp\Comp_Stats.csv"
}
}
Until ($Ans -eq 3)
Exit

}
}
Until ($Ans -eq 3)
