
    $computers = Get-Content "\\XLWUW-DJPVV1\C$\Users\1274873341C\Desktop\Desktop\PS_Scripts\_PSAppDeployToolkit_v3.6.8\Flash_Update\Toolkit\Flash_Targets.txt"
    $sourcefolder = "\\XLWUW-DJPVV1\C$\Users\1274873341C\Desktop\Desktop\PS_Scripts\_PSAppDeployToolkit_v3.6.8\Flash_Update"
    $destinationFolder = "\\$comp\c$\TEMP\Adobe_Flash"


foreach ($comp in $computers) {

    Try
        {            
           
            
         If (!(Test-Path -path $destinationFolder))
            
            {                        
                New-Item $destinationFolder -Type Directory -Force
            }
                Copy-Item -Path "$sourcefolder\*" -Destination $destinationFolder -Recurse -Force
                
                Invoke-Command -ScriptBlock {& cmd.exe /c "$destinationFolder\Toolkit\Deploy-Application.exe /s /v /qn" }
                
                #Invoke-Command {Start-Process -FilePath "C:\TEMP\Adobe_Flash\Toolkit\Deploy-Application.ps1" }

                #Invoke-Command -FilePath "($destinationFolder)\Toolkit\Deploy-Application.ps1"   
        }
    
    Catch
        {
            Add-Content "\\XLWUW-DJPVV1\C$\Users\1274873341C\Desktop\Desktop\PS_Scripts\_PSAppDeployToolkit_v3.6.8\Flash_Update\Toolkit\Failed-Computers.txt" $comp
        }
 
}
    
 

    # Invoke-Command -ScriptBlock {& cmd.exe /c "\\XLWUW-DJPVV1\C$\Users\1274873341C\Desktop\Desktop\PS_Scripts\_PSAppDeployToolkit_v3.6.8\Flash_Update\Toolkit\Deploy-Application.exe /s /v /qn" } 
 
    #\\XLWUW-DJPVV1\C$\Users\1274873341C\Desktop\Desktop\PS_Scripts\_PSAppDeployToolkit_v3.6.8\Flash_Update\Toolkit\Deploy-Application.ps1
    
    #-Credential $creds -Authentication Credssp

