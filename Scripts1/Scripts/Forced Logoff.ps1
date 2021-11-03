#-----------------------------------------------------------------------------------------#
#                                  Written by SrA David Roberson                          #
#                                  Tyndall AFB, Panama City, FL                           #
#                                     Created 08 August 2014                              #
#-----------------------------------------------------------------------------------------#

$Computers = Read-Host "Computer Name"
$Message = "You have interrupted a required system maintenance operation.  
                         You will be logged off in 2 minutes."
$Caption = "325 CS/SCOO Cyberspace Operations"

ForEach ($Computer in $Computers)
    {
    Get-TSSession -State Active -ComputerName $Computer | Send-TSMessage -Caption $Caption -Text $Message 
    Start-Sleep -Seconds 120
    Get-TSSession -State Active -ComputerName $Computer | Stop-TSSession -Force
    }