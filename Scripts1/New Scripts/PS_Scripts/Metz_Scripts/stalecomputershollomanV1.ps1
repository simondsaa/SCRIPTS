#########################################################################
#Script Name:stalecomputershickam.ps1
#Authored by: Turner S. Marshall
#             Chris Steele
#Date Created: 25 Jan 18
#
#Date Changed:
#Purpose: This Script is made to pull Holloman AFB computers and all 
#         relevant information.
#########################################################################
param($defaultlog = "$($env:userprofile)\Documents\computer_reporthollo.csv",
     $staledate = 90)
  
  
 #format date
 $stale_date = [DateTime]::Today.AddDays(-$staledate)
 #delete results if already exist
 If ($(Try { Test-Path $defaultlog} Catch { $false })){Remove-Item $defaultlog -force}
  
 #region create hashtables
  
 #this is hashtable used to populate a calculated property to determine if the account is stale
 $hash_isComputerStale = @{Name="Stale";
     Expression={if(($_.LastLogonTimeStamp -lt $stale_date.ToFileTimeUTC() -or $_.LastLogonTimeStamp -notlike "*") `
         -and ($_.pwdlastset -lt $stale_date.ToFileTimeUTC() -or $_.pwdlastset -eq 0) `
         -and ($_.enabled -eq $true) -and ($_.whencreated -lt $stale_date) `
         -and ($_.IPv4Address -eq $null) -and ($_.OperatingSystem -like "Windows*") `
         -and (!($_.serviceprincipalnames -like "*MSClusterVirtualServer*"))){$True}else{$False}}}
  
 #this hashtable is used to create a calculated property that converts pwdlastset
 $hash_pwdLastSet = @{Name="pwdLastSet";
     Expression={([datetime]::FromFileTime($_.pwdLastSet))}}
  
 #this hashtable is used to create a calculated property that converts lastlogontimestamp
 $hash_lastLogonTimestamp = @{Name="LastLogonTimeStamp";
     Expression={([datetime]::FromFileTime($_.LastLogonTimeStamp))}}
  
 #this hashtable is used to create a calculated property to display domain of the computer
 $hash_domain = @{Name="Domain";
     Expression={$domain}}
  
 #endregion
  
 foreach($domain in (get-adforest).domains){
     # Put your base DN between the quotes right below here
     get-adcomputer -SearchBase "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" -filter {isCriticalSystemObject -eq $False} `
         -properties PwdLastSet,whencreated,SamAccountName,LastLogonTimeStamp,
             Enabled,IPv4Address,operatingsystem,serviceprincipalnames `
         -server $domain | `
             select $hash_domain,SamAccountName,enabled,operatingsystem,IPv4Address,`
                 $hash_isComputerStale,$hash_pwdLastSet,$hash_lastLogonTimestamp | `
             export-csv $defaultlog -append -NoTypeInformation
 }
  
 $results = import-csv $defaultlog
 $stale =  $results | group-object stale | select name, count
 $disabled = $results | group-object enabled | select name, count
  
  
 Write-Host "Found $(($stale | where name -eq $true).count) stale computers"
 Write-Host "Found $(($disabled | where name -eq $false).count) disabled computers"
 Write-Host "Found $($results.count) total computers"
