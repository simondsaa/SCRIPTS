Write-Host "Verifying Media is up to date"
Sleep 8
$currentDir = pwd
$requiredDiskDate = [DATETIME] "05/31/2017"
$driveInfo = Get-CimInstance win32_LogicalDisk
foreach ($drive in $driveInfo) {
    if ($drive.DriveType -eq 5) {
        Set-Location $drive.DeviceID
        $searchCD = Get-ChildItem
        foreach ($item in $searchCD) {
            if (($item.name) -like 'boot') {
                $diskDate = $item.LastWriteTime
                if ($diskDate -lt $requiredDiskDate) {
                    [System.Windows.Forms.MessageBox]::Show("The Media is outdated.  Please create updated media and try again.  Thank you.", "                                                            Outdated Media", 0)
                    THROW 'ERROR'
                    }
                }
            }
                    
        }
    }
cd "$currentDir"