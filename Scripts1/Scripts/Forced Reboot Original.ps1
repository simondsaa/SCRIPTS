<#-----------------------------------------------------------------------------------------#
 #                                  Written by SrA David Roberson                          #
 #                                  Tyndall AFB, Panama City, FL                           #
 #                                     Created 08 August 2014                              #
 #-----------------------------------------------------------------------------------------#>

$Comps = Get-Content C:\work\TEST.txt
#$Comps = Read-Host "Computer Name"

ForEach ($Comp in $Comps)
    {
        If (Test-Connection $Comp -quiet -BufferSize 64 -Ea 0 -count 5)
            {
                #Get-TSSession -State Active -ComputerName $Computer | Send-TSMessage -Caption $Caption -Text $Message -ErrorAction Continue
                Shutdown /m \\$Comp /f /r /t 600 /c "Your computer dumb. It need to sleep. Night Night."
                #.\pstools\psshutdown.exe \\$Comp -m $Message -f -r -t 900
            }

            Else
                {
                    $result = "$Comp is not accessible."
                    $result | Out-File -Verbose C:\work\Reboot_Again_Failed.txt -Append
                }  
    }    

$Message = "A required update is being installed on your computer.  
Please save all of your work.  Your computer will reboot in 15 minutes.  

Please click OK to close this window."
$Caption = "325 CS/SCOO Cyberspace Operations"

