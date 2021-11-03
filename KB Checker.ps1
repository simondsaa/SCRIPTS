## Create output directory on executing user's desktop 
 $date = get-date -Format MMddyyyy 
 $rootPath = $env:USERPROFILE + '\Desktop\CISA' 
 $fullPath = $rootPath + '\' + $date 
 
 $UpdateList = $fullPath + '\ALERT.csv' 

 Write-Host -ForegroundColor White ("INFORMATION: DOES OUTPUT PATH EXIST?") 
 if (Test-Path $rootPath) 
     { 
         if (test-path $fullPath) 
             { 
                 Write-Host -ForegroundColor White ("INFORMATION: YES, OUTPUT PATH EXISTS") 
             } 
         else 
             { 
                 Write-Host -ForegroundColor White ("INFORMATION: NO, OUTPUT PATH DOES NOT EXIST") 
                 Write-Host -ForegroundColor White ("INFORMATION: CREATING OUTPUT DIRECTORY $fullPath") 
                 mkdir $fullPath 
             } 
     } 
 else 
     { 
         Write-Host -ForegroundColor White ("INFORMATION: NO, OUTPUT PATH DOES NOT EXIST") 
         Write-Host -ForegroundColor White ("INFORMATION: CREATING OUTPUT DIRECTORY $fullPath") 
         mkdir $fullPath 
     } 
 
 
 
 
 ## Get all DCs in forest 
 Write-Host -ForegroundColor White ("INFORMATION: Getting list of all DCs in $(Get-ADForest)") 
 $allDCs = $((Get-ADForest).Domains | %{ Get-ADDomainController -Filter * -Server $_ }) | select hostname 
 Write-Host -ForegroundColor White ("INFORMATION: List contains $($((Get-ADForest).Domains | %{ Get-ADDomainController -Filter * -Server $_ }).hostname.count) DCs") 
 
 
 ## Foreach DC, get Component Based Servicing provided updates and MSI installed updates. Then dump to a common CSV 
 $allDCs | % { 
     $TLS = $True 
     $DC = $_.Hostname 
 
 
      Write-Host -ForegroundColor White ("INFORMATION: Testing Secure WinRM...") 
       try{ 
         Test-WSMan -ComputerName $DC -UseSSL -ErrorAction Stop 
     } 
      Catch{ 
        $TLS = $false 
     } 
 
 
     If ($TLS -eq $True){ 
 
 
         Write-Host -ForegroundColor White ("INFORMATION: Using secure WinRM for $DC") 
 
 
         $OS = Invoke-Command -ComputerName $DC -UseSSL -ScriptBlock { $(Get-WmiObject -Class Win32_OperatingSystem).caption} 
 
 
         Write-Host -ForegroundColor White ("INFORMATION: Getting updates for $DC") 
         Write-Host -ForegroundColor White ("INFORMATION: CBS Updates...") 
 
 
         $CBSs = Invoke-Command -ComputerName $DC -UseSSL -ScriptBlock { 
 
 
             $CBSArray = @() 
 
 
             Get-WmiObject Win32_quickfixengineering | Select-Object * | ForEach-Object{ 
                 $Hotfix = $_.HotFixID 
                 $Type = "CBS" 
                 $OS = $(Get-WmiObject -Class Win32_OperatingSystem).caption 
                 $DC = $([System.Net.Dns]::GetHostByName(($env:computerName))).Hostname 
 
 
                 $result = switch ($Hotfix){ 
                     KB4571729 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4571719 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4571736 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4571702 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4571703 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
            } 
                     KB4571723 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                   $CBSArray += $CBSObj 
            } 
                    KB4571694 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4565349 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                $CBSArray += $CBSObj 
             } 
                     KB4565351 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4566782 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4577051 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4577038 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4577066 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4577015 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4577069 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4574727 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4577062 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4571744 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4571756 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4571748 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                 #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                 $CBSArray += $CBSObj 
             } 
                     KB4570333 { 
                 $CBSObj = New-Object -TypeName psobject 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "DomainController" -Value $DC 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value $OS 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Update" -Value $Hotfix 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Type" -Value $Type 
                 $CBSObj | Add-Member -MemberType NoteProperty -Name "Compliance" -Value $True 
                #$CBSObj | Export-Csv -path $UpdateList -Append -NoTypeInformation 
                $CBSArray += $CBSObj 
             }  
 
