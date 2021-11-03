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
                    Write-Progress -Activity "Applying MSUs.. $i of $count" `
                                   -Status "Installing $KB... -- $([Math]::round($i/$count*100))% complete" `
                                   -PercentComplete ($i/$count*100)
                }
                else {
                    $i++
                    Write-Progress -Activity "Applying MSUs.. $i of $count" `
                                   -Status "$KB is installed... -- $([Math]::round($i/$count*100))% complete" `
                                   -PercentComplete ($i/$count*100)
                }
            }
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