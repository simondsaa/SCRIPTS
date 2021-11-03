Write-Host "Verifying Secure Boot is enabled and configured correctly"
sleep 8
$confirm = Confirm-SecureBootUEFI -ErrorVariable confirmError
$setupMode = (Get-SecureBootUEFI -Name SetupMode).bytes
$secureBoot = (Get-SecureBootUEFI -Name SecureBoot).bytes
if (($confirmError -ne $null) -or ($confirm -ne $true) -or ($setupMode -ne 0) -or ($secureBoot -ne 1)) {
    [System.Windows.Forms.MessageBox]::Show("`n`n`n`n`n`n`n`n`nThis system has not been properly configured.  Please ensure that the follow has been configured and try again.  Thank you`n`nSecure Boot is enabled`nLegacy Support is Disabled`nSetup Mode is correct", "BIOS/UEFI CONFIGURATION ERROR", 0)
    THROW 'ERROR'
    }