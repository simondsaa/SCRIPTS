
  
    
    $sourcefile = "\\52xlwu-ps-001p\c$\Users\1274873341.adm\Desktop\Program Scripts\Flash\Installers\*"
    $destinationFolder = "C:\TEMP\Flash"
    
#Uninstalls any version of Flash currently installed  
   
  <#  $App = Get-WmiObject Win32_Product | Where {$_.Name -like "*Flash*"}
    $App.Uninstall() 
    #>

#  Checks for destination working directory and creates it if necessary  
     If (!(Test-Path -path $destinationFolder))
            
            {
                        
                New-Item $destinationFolder -Type Directory -Force
            }

                Copy-Item -Path $sourcefile -Destination $destinationFolder 

                Stop-Process -Force -ProcessName iexplore, chrome, firefox, *flash*

                Invoke-Command  -ScriptBlock {cmd /c "start /wait msiexec.exe /i C:\TEMP\Flash\install_plugin.msi /qn"}
                Invoke-Command  -ScriptBlock {cmd /c "start /wait msiexec.exe /i C:\TEMP\Flash\install_active_x.msi /qn"}       