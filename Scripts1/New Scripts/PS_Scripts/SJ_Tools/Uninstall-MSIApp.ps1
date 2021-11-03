<#
.EXAMPLES 
Uninstall-MSIApp - ApplicationName "Adobe Flash" -Switches "/qn /norestart"
Uninstall-MSIApp - ApplicationName "Firefox" -Switches "/quiet /forcerestart"
#>
function Uninstall-MSIApp {

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$ApplicationName,
    [string]$Switches
)

    $ErrorActionPreference = 'SilentlyContinue'
    $UninstallPath = Get-ChildItem HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall -Recurse
    $UninstallPath += Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ -Recurse
    $SearchAppName = "*" + $ApplicationName + "*"
    $MSIPath = $Env:windir + "\system32\msiexec.exe"
    foreach ($item in $UninstallPath) {
        $tempitem = $item.name -split "\\"
        if ($tempitem[2] -eq "Microsoft") {
            $item = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $item.PSChildName
        }
        else {
            $item = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" + $item.PSChildName
        }
        if ((Test-Path $item) -eq $true) {
            $itemname = Get-ItemProperty -Path $item
            if ($itemname.displayname -like $SearchAppName) {
                $tempitem = $itemname.UninstallString -split " "
                if ($tempitem[0] -like "Msiexec.exe") {
                    Write-Host "Uninstall"$itemname.DisplayName"....." -NoNewline
                    $Parameters = "/x " + $itemname.PSChildName + [char]32 + $Switches
                    $ErrCode = (Start-Process -FilePath $MSIPath -ArgumentList $Parameters -Wait -Passthru).ExitCode
                    if (($ErrCode -eq 0) -or ($ErrCode -eq 3010) -or ($ErrCode -eq 1605)) {
                        Write-Host "Success" -ForegroundColor Green
                    }
                    else {
                        Write-Host "Failed with error code "$ErrCode -ForegroundColor Red
                    }
                }
            }         
        }
    }
}

Uninstall-MSIApp -ApplicationName "Microsoft SQL Server 2005" -Switches "/qn /norestart"