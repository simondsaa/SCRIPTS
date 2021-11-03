<#
Link to Guide:  https://www.danielengberg.com/hp-bios-configuration-utility-sccm/

.DESCRIPTION
Sets HP UEFI configuration
    
.NOTES

Author: Daniel Classon
Version: 1.1
Date: 2018-10-31
    
.EXAMPLE
.\Set-HPConfiguration.ps1 -Enable TPM

.DISCLAIMER
All scripts and other powershell references are offered AS IS with no warranty.
These script and functions are tested in my environment and it is recommended that you test these scripts in a test environment before using in your production environment.
#>

Param(
    [Parameter(Mandatory=$False)]
    [string]$Enable,
    [Parameter(Mandatory=$False)]
    [string]$Disable,
    [Parameter(Mandatory=$False)]
    [string]$Configure,
    [Parameter(Mandatory=$False)]
    [string]$PasswordFile = "$PSScriptRootpwd.bin",
    [Parameter(Mandatory=$False)]
    [string]$PasswordFile2 = "$PSScriptRootpwd2.bin"

)

Begin {

    switch ($Configure)
    {
        'ThunderboltSecurity' {$ConfigFile="$PSScriptRootConfigure_ThunderboltSecurity.txt"}
        'VideoMemory' {$ConfigFile="$PSScriptRootConfigure_VideoMemory.txt"}
        Default {}
    }

    switch ($Enable)
    {
        'SecureBoot' {$ConfigFile="$PSScriptRootEnable_SecureBoot.txt"}
        'TPM' {$ConfigFile="$PSScriptRootEnable_TPM.txt"}
        'Virtualization' {$ConfigFile="$PSScriptRootEnable_Virtualization.txt"}
        'WLANSwitching' {$ConfigFile="$PSScriptRootEnable_WLANSwitching.txt"}
        Default {}
    }
        switch ($Disable)
    {
        'Virtualization' {$ConfigFile="$PSScriptRootDisable_Virtualization.txt"}
        Default {}
    }

}
Process {
    $process = Start-Process -FilePath "$PSScriptRootBiosConfigUtility64.exe" -ArgumentList "`"/npwdfile:$PasswordFile`"", "`"/set:$ConfigFile`"", "/log" -Wait -PassThru

    #If a password is configured, enter it
    if ($process.ExitCode -eq 10) {
        try {
            $process = Start-Process -FilePath "$PSScriptRootBiosConfigUtility64.exe" -ArgumentList "`"/cpwdfile:$PasswordFile`"", "`"/set:$ConfigFile`"", "/log" -wait -PassThru
        }
        catch {
            $process = Start-Process -FilePath "$PSScriptRootBiosConfigUtility64.exe" -ArgumentList "`"/cpwdfile:$PasswordFile2`"", "`"/set:$ConfigFile`"", "/log" -wait -PassThru
        }
    }
}  
End {
}           