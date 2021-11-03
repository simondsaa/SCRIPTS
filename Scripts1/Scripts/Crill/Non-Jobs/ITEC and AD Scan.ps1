"Importing ADPE and AD Info"

# Create blank line
$nl = [Environment]::newline


$ADPE = import-csv "C:\Users\1252862141.adm\Desktop\ITEC Scan\ALL Systems.csv"
$AD = import-csv "C:\Users\1252862141.adm\Desktop\ITEC Scan\SCMM List.csv"
$ADInfo = Get-ADComputer -filter * -Searchbase "OU=Tyndall AFB, OU=AFCONUSEAST,OU=BASES,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" -Properties Name, LastLogonDate, LogonCount, whenCreated |select  Name, LastLogonDate, LogonCount, whenCreated

"Complete"

$MatchedList = @()
$Unmatchedlist = @()
$counter = $null

Foreach ($obj in $adpe) {
    $MatchedList += $obj |where {$AD.SerialNumber -contains $_.SerialNumber}
    $Unmatchedlist +=  $obj |where {$AD.SerialNumber -notcontains $_.SerialNumber}
    $counter++
    Write-Progress -activity “Building Lists” -status “Progress : $counter / $($adpe.count)” -PercentComplete (($counter / $adpe.count)*100)
    }

"Matched $($MatchedList.count)"
"Umatched $($Unmatchedlist.count)"


"Attempting to find more matches"



$counter = $null
$CombinedList = @()
Foreach ($obj in $MatchedList) {
    $ADInfo = $ad |where {$_.Serial -eq $obj.SerialNumber}
    $converter = [PSCUstomObject]@{
        ComputerName = $adinfo.Target_ID
        Organization = $obj.OrgName
        Account = $obj.AccountNumber
        ITEC_InstallDate = $obj.InstallDate
        AD_LastLogon = $ADInfoScan.LastLogonDate
        AD_LogonCount = $ADInfoScan.LogonCount
        AD_CreationTime = $ADInfoScan.whenCreated
        SerialNumber = $ADInfo.Serial
        Building = $obj.Building
        Room = $obj.Room
    }

    $combinedList += $converter
    $counter++
    Write-Progress -activity “Building Lists” -status “Progress : $counter / $($matchedList.count)” -PercentComplete (($counter / $($matchedList.count))*100)
    }


"Information Compilation  Complete" 
 $Date = Get-Date -UFormat "%d-%b-%g %H%M" 

$CombinedList | Export-Csv "C:\Users\1252862141.adm\Desktop\ITEC Scan\MatchedList_$date.csv"
$UnmatchedList | Export-Csv "C:\Users\1252862141.adm\Desktop\ITEC Scan\UnmatchedList_$date.csv"
