$Computers = gc "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\HBSS\No_Agent_targets.txt"

foreach($strComputer in $computers)
{
    If((GWmi win32_operatingsystem -computername $strComputer).osarchitecture -eq "32-bit")
                    {
                    $executable = "Frminst.exe"
                    $switches = "/remove=agent"
                    $cmd = "\\$strComputer\c$\Program Files\McAfee\Common Framework\$executable $switches"
                    $ErrorActionPreference="SilentlyContinue"

                    $Ping = new-object system.net.networkinformation.ping
                    $reply = $ping.send($strComputer)

                        if ($reply.status -eq "Success")
                        {
                            Write-Host "$strComputer is online- 32-bit - Connecting now..." -ForegroundColor GREEN
   
                            Trap {write-Warning "There was an error connecting to the remote computer or creating the process"; continue}      
       
                            $wmi=([wmiclass]"\\$strComputer\root\cimv2:win32_process")   
    
                            if (!$wmi) {return}      
                            $remote=$wmi.Create($cmd)  
                            
                            $remote
                            if ($remote.returnvalue -eq 0) 
                            {     
                                Write-Host "Successfully launched on $strComputer" -ForegroundColor GREEN
                            } 
                            else 
                            {     
                                Write-Host "Failed to launch $cmd on $strComputer." -ForegroundColor RED 
                            } 
                        }
                        else
                        {
                            Write-Host "$strComputer - System offline" -ForegroundColor RED
                        }

                    }
    If((GWmi win32_operatingsystem -computername $strComputer).osarchitecture -eq "64-bit")
                    {
                    
                    $executable = "Frminst.exe"
                    $switches = "/forceuninstall"
                    $cmd = "\\$strComputer\c$\Program Files (x86)\McAfee\Common Framework\$executable $switches"
                    $ErrorActionPreference="SilentlyContinue"

                    $Ping = new-object system.net.networkinformation.ping
                    $reply = $ping.send($strComputer)

                        if ($reply.status -eq "Success")
                        {
                            Write-Host "$strComputer is online- 64-bit - Connecting now..." -ForegroundColor GREEN
   
                            Trap {write-Warning "There was an error connecting to the remote computer or creating the process"; continue}      
       
                            $wmi=([wmiclass]"\\$strComputer\root\cimv2:win32_process")   
    
                            if (!$wmi) {return}      
                            $remote=$wmi.Create($cmd) 
 
    
                            if ($remote.returnvalue -eq 0) 
                            {     
                                Write-Host "Successfully launched on $strComputer" -ForegroundColor GREEN
                            } 
                            else 
                            {     
                                Write-Host "Failed to launch $cmd on $strComputer" -ForegroundColor RED 
                            } 
                        }
                        else
                        {
                            Write-Host "$strComputer - System offline" -ForegroundColor RED
                        }

                    }

}