<#
.AUTHOR
SSgt Gutierrez

.SYNOPSIS
This script will restart computer(s) with message.

.EXAMPLE
Reboot-Computer -ComputerName Client1 -Seconds 3600
Reboot-Computer -ComputerName Client1,Client2 -Seconds 7200

.NOTES
The Seconds parameter is based on seconds/hour. 

#>

function Reboot-Computer {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [string[]]$ComputerName,
        [string]$Seconds
        )

    Begin{
        $ErrorActionPreference = "SilentlyContinue"
        $time = ((Get-Date).ToShortTimeString())
        $date = ((Get-Date).ToLongDateString())
        $Minute = $Seconds/60 -as [string]
        $Hour = $Minute/60 -as [string]
    }

    Process{
        foreach($Computer in $ComputerName) {
           if(Test-Connection -ComputerName $Computer -Count 1) {
                $username = (Get-WmiObject -Class win32_computersystem -ComputerName $ComputerName).username
                $user= $username.split("\")[1]
                $u = $user.split(".")[0]
                $type = "SamAccountName"
                $DN = Get-ADUser -SearchScope Subtree -SearchBase "OU=Seymour Johnson AFB Users,OU=Seymour Johnson AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" -Filter {$type -like $u} -Properties DisplayName
                $DName = $DN.DisplayName
                $comment = "Attention $DName!!! Your system ($Computer) will restart in $Hour hour. Please save all your work to prevent data loss. Scheduled reboot at $time on $date."
                shutdown.exe /r /m \\$Computer /t $Seconds /f /c $comment
           }
        }
    }

    End{ 
    }
}