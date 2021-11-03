Function Install-Software
{
    <#  
    .SYNOPSIS
        Installs a program using a MSI file on remote computer silently.

    .DESCRIPTION
        Installs a program using a MSI file on remote computer silently. Requires the target computername, admin shares enabled and
        appropriate rights to install software on target comptuer.

    .PARAMETER ComputerName
        The name of the target computer

    .PARAMETER InstallPath
        The source path for the msi installer file to be used.

    .EXAMPLE
        Install-Software -ComputerName "THATPC" -InstallPath "C:\Scripts\Adobe\FlashPackage\install_flash_player_20_plugin.msi"

        Installs the flash player NPAPI plugin on remote computer named THATPC.

    .EXAMPLE
        Get-Content .\Computers.txt | Install-Software -InstallPath "C:\Scripts\Adobe\FlashPackage\install_flash_player_20_plugin.msi"

        Installs the NPAPI flash plugin software on remote computers listed in the computers.txt file.
        Computers.txt in this example has one computername per line.        

#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        $ComputerName,
        [Parameter(Mandatory=$True)]
        [string]$InstallPath
        )
    Copy-Item -Path $InstallPath -Destination "\\$ComputerName\c$\"
    $InstallFileName = ($InstallPath -split '\\')[-1]
    $MSIInstallPath = "\\$ComputerName\c$\$($InstallFileName)"
$returnval = ([WMICLASS]"\\$computerName\ROOT\CIMV2:win32_process").Create("msiexec `/i$MSIInstallPath `/norestart `/qn")
$returnval
}