$directory = "C:\work\bowling.txt"
$csvFiles = get-childitem $directory -filter *.csv

$results = @();
Foreach ($csv in $csvFiles) {
    $results += import-csv $csv
    $i++
    Write-Host "." -ForegroundColor Cyan
    Write-Progress -activity “Combining Information” -status “Status: $i/$($csvfiles.count)” -PercentComplete (($i / $csvFiles.count)*100)
        }

$i = $null

$Active_List = @()
$Active_List +=  $Results | where {$_.Remoting -eq "Active"} | select ComputerName, Remoting
$Inactive_List = @()
$Inactive_List +=  $Results | where {$_.Remoting -eq "InActive"} | select ComputerName, Remoting
$Repaired_List = @()

Foreach ($comp in $Inactive_List.ComputerName) {
psexec \\$comp /accepteula -s -d c:\windows\system32\winrm.cmd quickconfig -quiet
Enter-pssession $comp
$RemoteObj = New-Object -TypeName PSobject

    $WinRMTest =  Get-Service winrm
    If ($WinRMTest.Status -eq "Running") {
    $WinRM = "Active"
    }
    ELSE {
    $WINRM = "Inactive"
         }

$RemoteObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $comp
$RemoteObj | Add-Member -MemberType NoteProperty -Name Remoting -Value $WinRM

$Repaired_List +=  $RemoteObj
Exit-PSSession

}

get-pssession | remove-pssession 

cls
$Active_List += $Repaired_List
$ActiveCount = $Active_list | where {$_.Remoting -eq "Active"}

Write-Host "Pre-Fix Active: $(($Active_List).count)"  -ForegroundColor Green

Write-Host "Attemped to fix: $(($Inactive_List).count)"  -ForegroundColor Green

Write-Host "SUCCESSFULLY FIXED: $(($ActiveCount).count)"  -ForegroundColor Green



