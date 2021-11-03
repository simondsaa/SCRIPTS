<#

.SYNOPSIS
Get BIOS version from computers.

.PARAMETER ComputerName
Specifies the computers to query.

.PARAMETER IncludeNonResonding
Optional switch to include nonresponding computers.

.INPUTS
None. You cannot pipe objects.

.OUTPUTS
System.Object

.EXAMPLE
.\Get-BiosVersion

.EXAMPLE
.\Get-BiosVersion -ComputerName PC01,PC02,PC03

.EXAMPLE
.\Get-BiosVersion (Get-Content C:\computers.txt) -ErrorAction SilentlyContinue

.EXAMPLE
.\Get-BiosVersion (Get-Content C:\computers.txt) -Verbose -IncludeNonResponding |
Export-Csv BiosVersion.csv -NoTypeInformation

.NOTES
Author: Matthew D. Daugherty
Date Modified: 2 August 2020

#>

[CmdletBinding()]
param (

    [Parameter()]
    [string[]]
    $ComputerName = $env:COMPUTERNAME,

    [Parameter()]
    [switch]
    $IncludeNonResponding
)


# Scriptblock for Invoke-Command
$InvokeCommandScriptBlock = {

    $VerbosePreference = $Using:VerbosePreference

    Write-Verbose "Getting BIOS version on $env:COMPUTERNAME."

    $BIOS = Get-CimInstance -ClassName Win32_BIOS -Verbose:$false

    [PSCustomObject]@{

        ComputerName = $env:COMPUTERNAME
        SerialNumber = $BIOS.SerialNumber
        Manufacturer = $BIOS.Manufacturer
        Version = $BIOS.Name
    }
}

# Parameters for Invoke-Command
$InvokeCommandParams = @{

    ComputerName = $ComputerName
    ScriptBlock = $InvokeCommandScriptBlock
    ErrorAction = $ErrorActionPreference
}

switch ($IncludeNonResponding.IsPresent) {

    'True' {

        $InvokeCommandParams.Add('ErrorVariable','NonResponding')

        Invoke-Command @InvokeCommandParams | 
        Select-Object -Property *, ErrorId -ExcludeProperty PSComputerName, PSShowComputerName, RunspaceId

        if ($NonResponding) {

            foreach ($Computer in $NonResponding) {

                [PSCustomObject]@{

                    ComputerName = $Computer.TargetObject.ToUpper()
                    SerialNumber = $null
                    Manufacturer = $null
                    Version = $null
                    ErrorId = $Computer.FullyQualifiedErrorId
                }
            }
        }
    }
    'False' {

        Invoke-Command @InvokeCommandParams | 
        Select-Object -Property * -ExcludeProperty PSComputerName, PSShowComputerName, RunspaceId
    }
}
