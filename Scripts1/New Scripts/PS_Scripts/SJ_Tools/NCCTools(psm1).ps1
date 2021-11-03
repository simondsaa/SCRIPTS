function Delete-DRAComputer {
    <#
    .SYNOPSIS
    Deletes computer object(s) from DRA.
    .PARAMETER ComputerName
    The name of one or more computers.
    .PARAMETER LogErrors
    Log failed computer names to a text file.
    .PARAMETER ErrorLog
    The file name to log computer names to - defaults to C:\Users\$env:USERNAME\Desktop\Delete-DRAComputer_errors.txt.
    #>
    param(
        [cmdletbinding()]
        [parameter(ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True,
                   Mandatory=$True)]
        [string[]]$ComputerName,
        [string]$errorlog = "C:\Users\$env:USERNAME\Desktop\Delete-DRAComputer_errors.txt",
        [switch]$LogErrors
        )
    Begin {
        $date = Get-Date -Format F
        $i = 0
        $total = $ComputerName
    }
    Process {
        foreach ($computer in $ComputerName) {            
            try {
                $i++
                Remove-DRAComputer -Domain $env:USERDNSDOMAIN `
                                   -DRAHostServer localhost `
                                   -DRAHostPort 11192 `
                                   -DRARestServer seymourjohnson.dra.us.af.mil `
                                   -DRARestPort 8755 `
                                   -Identifier $("CN=$computer,OU=Seymour Johnson AFB Computers,OU=Seymour Johnson AFB,OU=AFCONUSEAST,OU=BASES,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL") `
                                   -Force `
                                   -ErrorAction Stop
                Write-Progress -Activity "Removing computer objects from ADUC --- $i of $total" `
                               -Status "Removing $computer from ADUC --- $([Math]::round($i/$total*100))% complete" `
                               -PercentComplete ($i/$total*100)
                Write-Output "Removing $computer on $date" >> "C:\Users\$env:USERNAME\Desktop\RemovedComputers.txt"
            }
            catch {
                if ($LogErrors) {
                    $i--
                    "Failed to delete $computer" | 
                    Out-File $errorlog -Append
                }
                Write-Warning "$computer failed!"
            }
        }
    }
    End {
        Write-Host "Successfully deleted $i computer object(s) from DRA -- $date" -ForegroundColor Green
    }
}

Delete-DRAComputer -ComputerName (Get-Content -Path "C:\Users\$env:USERNAME\Desktop\removecomputerlist.txt") -LogErrors
########################################################################################################################
function Send-Popup {
    <#
    .SYNOPSIS
    Sends pop-up to computer(s).
    .PARAMETER ComputerName
    The name of one or more computers.
    .PARAMETER LogErrors
    Log failed computer names to a text file.
    .PARAMETER ErrorLog
    The file name to log computer names to - defaults to C:\Users\$env:USERNAME\Desktop\popup_failed.txt.txt.
    #>
    param(
        [cmdletbinding()]
        [parameter(ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True,
                   Mandatory=$True)]
        [string[]]$ComputerName,
        [string]$errorlog = "C:\Users\$env:USERNAME\Desktop\popup_failed.txt",
        [switch]$LogErrors
        )
    Begin {
        $date = Get-Date -Format F
        $i = 0
        $message = "Per MTO 2017-125-001, the AF Chief Information Officer mandated transition to the Microsoft Windows 10 Operating System on the Air Force Network. IAW this MTO, 4 FW/CC mandates that Win 7 operating systems be upgraded by 01 Feb 18."
    }
    Process {
        foreach($Computer in $ComputerName) {
            try {
                $i++
                Invoke-WmiMethod -Path Win32_Process `
                                 -Name Create `
                                 -ArgumentList "msg * /time:3600 $message" `
                                 -ComputerName $Computer `
                                 -ErrorAction Stop
                Write-Output "Sent pop up to $Computer on $date" >> "C:\Users\$env:USERNAME\Desktop\popup_successfull.txt"
            }
            catch {
                if ($LogErrors) {
                    $i--
                    "Failed to send pop up on $Computer on $date" | 
                    Out-File $errorlog -Append
                }
                Write-Warning "$Computer failed!"
            }  
        }
    }
    End {
        Write-Host "Sent pop-up to $i system(s) -- $date" -ForegroundColor Green
    }
}

Send-Popup -ComputerName (Get-Content "C:\Users\$env:USERNAME\Desktop\popuplist.txt") -Logerrors
########################################################################################################################
function Remove-SoftCerts {
    <#
    .SYNOPSIS
    Removes soft certs (.p12 & .pfx) from computer(s).
    .PARAMETER ComputerName
    The name of one or more computers.
    .PARAMETER LogErrors
    Log failed computer names to a text file.
    .PARAMETER ErrorLog
    The file name to log computer names to - defaults to C:\Users\$env:USERNAME\Desktop\Remove-SoftCerts_Error.txt.
    #>
    param(
        [cmdletbinding()]
        [parameter(ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True,
                   Mandatory=$True)]
        [string[]]$ComputerName,
        [string]$errorlog = "C:\Users\$env:USERNAME\Desktop\Remove-SoftCerts_error.txt",
        [switch]$LogErrors
        )
    Begin {
        $date = Get-Date -Format F
        $i = 0
    }
    Process {
        foreach($computer in $ComputerName) {
            try {
                $i++
                Invoke-Command -ComputerName $computer -ScriptBlock {
                    if (Get-ChildItem -Path c:\ -Include *.p12,*.pfx -File -Recurse -ErrorAction SilentlyContinue) {
                        Get-ChildItem -Path c:\ -Include *.p12,*.pfx -File -Recurse -ErrorAction SilentlyContinue |
                        foreach $_ {
                            $file = ($_).FullName
                            Remove-Item -Path $file -Force
                        }
                    }
                }
                Write-Output "Removed soft certs from $computer on $date" >> "C:\Users\$env:USERNAME\Desktop\SoftCertsRemoved.txt"
            }
            catch {
                if ($LogErrors) {
                    $i--
                    "Failed to remove soft certs from $computer on $date" | 
                    Out-File $errorlog -Append
                }
                Write-Warning "$computer failed!"
            }  
        }
    }
    End {
        Write-Host "Ran soft cert removal script on $i system(s) -- $date" -ForegroundColor Green
    }
}

Remove-SoftCerts -ComputerName vkagl-02240h -LogErrors
########################################################################################################################
function Install-SecurityUpdates {
    <#
    .Synopsis
       Installs all .msu files in a specific directory.
    .EXAMPLE
       Install-SecurityUpdates
    .REQUIREMENTS
       The path must have .msu files for this script to work. This script must be run locally and not remotely.
    #>
    $i = 0
    $vb = New-Object -ComObject wscript.shell
    $filepath = New-Object System.Windows.Forms.FolderBrowserDialog
    $result = $filepath.ShowDialog() 
    if ($result -eq 'OK') {
        $fp = $filepath.SelectedPath
        $Dir = (Get-Item -Path $($fp)  -Verbose).FullName
        $MSUs = ls  -Path $Dir -Filter *.msu
        $count = $MSUs.count
        $answer = $vb.popup("About to install $($count) update(s) on $($env:COMPUTERNAME). Do you want to proceed?",0,"MSU Installer",4)
        if ($answer -eq 6) {
            foreach ($MSU in $MSUs){
                if ($MSU.Name -like "WinSec-*") {
                    $update = $MSU.Name -split'-'
                    $KB = $update[6]
                }
                elseif ($MSU.Name -like "MS*") {
                    $update = $MSU.Name -split'-'
                    $KB = $update[2]
                }
                elseif ($MSU.Name -like "Windows*") {
                    $update = $MSU.Name -split'-'
                    $KB = $update[1]
                }
                else {
                    $update = $MSU.Name -split'-'
                    $KB = $update[2]   
                }
                $HotFix = Get-HotFix -Id $KB -ErrorAction SilentlyContinue
                if ($HotFix -eq $null) {
                    $i++
                    $InstallString = $Dir + "\" + $MSU.Name
                    wusa.exe $InstallString /quiet /norestart | Out-Null
                    Write-Progress -Activity "Applying security updates --- $i of $count..." `
                                   -Status "Installing $KB --- $([Math]::round($i/$count*100))% complete.." `
                                   -PercentComplete ($i/$count*100)
                }
                else {
                    $i++
                    Write-Progress -Activity "Applying security updates --- $i of $count..." `
                                   -Status "$KB is installed --- $([Math]::round($i/$count*100))% complete.." `
                                   -PercentComplete ($i/$count*100)
                    $i--
                }
            }
            Write-Host "Attempted to install $i update(s) on $env:COMPUTERNAME" -ForegroundColor Green
        }

        else {
            $vb.Popup("Script Ended!",0,"Bye!",0)
        }
    }
    else {
        Write-Error "User cancelled script."
    }
    $DeleteFolder = $vb.Popup("Do you wish to delete the folder?",0,"Remove Updates Folder",4)
    if ($DeleteFolder -eq 6) {
        Remove-Item -Path $fp -Force -Recurse
    }
    else {
        $vb.Popup("Script Ended!",0,"Bye!",0)
    }
}

Install-SecurtiyUpdates
########################################################################################################################
function Do-McAfee {
    <#
    .SYNOPSIS
    Can install/uninstall/update McAfee on target computer(s)
    .PARAMETER ComputerName
    The name of one or more computers.
    .PARAMETER LogErrors
    Log failed computer names to a text file.
    .PARAMETER ErrorLog
    The file name to log computer names to - defaults to C:\Users\$env:USERNAME\Desktop\Do-McAfee_errors.txt.
    .EXAMPLE
    Do-Mcafee -ComputerName computername -Uninstall -LogErrors
    Do-Mcafee -ComputerName computername1,computername2 -Install -LogErrors
    Do-Mcafee -ComputerName (Get-Content "C:\Users\$env:USERNAME\Desktop\McAfeeTargetList.txt") -Update -LogErrors
    #>
    param(
        [parameter(Mandatory=$True)]
        [String[]]$ComputerName,
        [String]$errorlog = "C:\Users\$env:USERNAME\Desktop\Do-McAfee_errors.txt",
        [Switch]$Uninstall,
        [Switch]$Install,
        [Switch]$Update,
        [switch]$LogErrors
    )


    Begin {
        $date = Get-Date -Format F
        $i = 0
        $rh = 0
        $total = $ComputerName.Count
    }
    Process {
        foreach ($computer in $ComputerName) {
            try {
                if($Uninstall) {
                    if(Test-Connection -ComputerName $computer -Quiet -Count 1) {
                        $status = "-Uninstall"
                        $rh++
                        $i++
                        Invoke-Command -ComputerName $computer -ErrorAction Stop -ScriptBlock {
                            Start -Wait `
                                  -FilePath 'C:\Program Files\McAfee\Agent\x86\FrmInst.exe' `
                                  -ArgumentList '/forceuninstall' `
                                  -ErrorAction Stop
                        }
                        Write-Progress -Activity "Uninstalling McAfee Agent --- $i of $total" `
                                       -Status "Uninstalling McAfee on $computer --- $([Math]::Round($i/$total*100))% complete" `
                                       -PercentComplete ($i/$total*100)
                    }
                    else {
                        $rh--
                        Write-Host "$computer is offline or non-existent" -ForegroundColor Red
                    }
                }

                if($Install) {
                    if(Test-Connection -ComputerName $computer -Quiet -Count 1) {
                        $status = "-Install"
                        $rh++
                        $i++
                        Copy-item -Path "\\vkag-fs-01pv\SeymourJohnson_4FW_SJ_ALL\sj_all_csa_info\Software Patches\HBSS & McAfee AV\FramePkg.exe" `
                                  -Destination "\\$computer\C$" `
                                  -Force
                        Invoke-Command -ComputerName $computer -ErrorAction Stop -ScriptBlock {
                            Start -Wait `
                                  -FilePath "C:\FramePkg.exe" `
                                  -ArgumentList '/install=agent /forceinstall' `
                                  -ErrorAction Stop
                            Remove-Item -Path "C:\FramePkg.exe" -Force
                        }
                        Write-Progress -Activity "Installing Framepkg --- $i of $total" `
                                       -Status "Installing McAfee on $computer --- $([Math]::Round($i/$total*100))% complete" `
                                       -PercentComplete ($i/$total*100)
                    } 
                    else {
                        $rh--
                        Write-Host "$computer is offline or non-existent" -ForegroundColor Red
                    }
                }


                if($Update) {
                    if (Test-Connection -ComputerName $computer -Quiet -Count 1) {
                        $status = "-Update"
                        $rh++
                        $i++
                        Invoke-Command -ComputerName $computer -ErrorAction Stop -ScriptBlock {
                            if(Test-Path -Path 'C:\Program Files (x86)\McAfee\VirusScan Enterprise\') {
                                Start -Wait `
                                      -FilePath 'C:\Program Files (x86)\McAfee\VirusScan Enterprise\mcupdate.exe' `
                                      -ErrorAction Stop
                            } 
                            else {
                                Start -Wait `
                                      -FilePath 'C:\Program Files\McAfee\VirusScan Enterprise\mcupdate.exe' `
                                      -ErrorAction Stop
                            }
                        }
                        Write-Progress -Activity "Updating McAfee --- $i of $total" `
                                       -Status "Updating security on $computer --- $([Math]::Round($i/$total*100))% complete"`
                                       -PercentComplete ($i/$total*100)
                    }
                    else {
                        $rh--
                        Write-Host "$computer is offline or non-existent" -ForegroundColor Red
                    }
                }
            }
            catch {
                if ($LogErrors) {
                    if ($Uninstall) {
                        "Failed to uninstall McAfee from $computer on $date" |
                        Out-File $errorlog -Append
                    }
                    if ($Install) {
                        "Failed to install McAfee from $computer on $date" |
                        Out-File $errorlog -Append
                    }
                    if ($Update) {
                        "Failed to update McAfee from $computer on $date" |
                        Out-File $errorlog -Append
                    }
                }
                $rh--
                Write-Warning "$computer failed!"
            }
        }
    }
    End {
        Write-Host "Successfully ran Do-McAfee $($status) on $($rh) computer(s) -- $($date)" -ForegroundColor Green
    }
}

Do-Mcafee -ComputerName (Get-Content -Path C:\Users\1460568001.adm\Desktop\rogues.txt) -Install -LogErrors
########################################################################################################################
<#
.SYNOPSIS
Get all SJ systems
#>
Get-ADComputer -SearchScope Subtree `
               -SearchBase "OU=Seymour Johnson AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" `
               -Filter {(ObjectCategory -eq "Computer") -And (OperatingSystem -like '*Windows*')} `
               -Properties * | 
Select-Object Name,
              OperatingSystem,
              OperatingSystemServicePack,
              OperatingSystemVersion,
              IPV4Address,
              @{n='ObjectLocation';e={$_.CanonicalName}},
              whenCreated,
              LastLogonDate,
              @{L='Organization';E={$_.o[0]}},
              Location,
              Enabled |
Sort-Object Name |
Export-CSV  "C:\Users\$env:USERNAME\Desktop\SJAllSystems.csv" -NoTypeInformation -Encoding UTF8

Get-ADComputer -SearchScope Subtree `
               -SearchBase "OU=Servers,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" `
               -Filter {(ObjectCategory -eq 'Computer') -and (OperatingSystem -like '*Windows*') -and (Name -like '*vkag*') -or (Name -like 'gsb*')} `
               -Properties *| 
Select-Object Name,
              OperatingSystem,
              OperatingSystemServicePack,
              OperatingSystemVersion,
              IPV4Address,
              @{n='ObjectLocation';e={$_.CanonicalName}},
              whenCreated,
              LastLogonDate,
              @{L='Organization';E={$_.o[0]}},
              Location,
              Enabled |
Sort-Object Name | 
Export-CSV  "C:\Users\$env:USERNAME\Desktop\SJAllSystems.csv" -NoTypeInformation -Encoding UTF8 -Append