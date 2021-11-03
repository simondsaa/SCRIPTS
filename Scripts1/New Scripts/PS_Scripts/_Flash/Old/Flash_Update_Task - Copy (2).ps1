
  
    
    $sourcefile_x64 = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Applications\Adobe Flash\Adobe Flash 21.0.0.242\*"
    $sourcefile_x86 = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Applications\Adobe Flash\Adobe Flash 21.0.0.242\*"
    $destinationFolder_x64 = "C:\TEMP\Flash_x64"
    $destinationFolder_x86 = "C:\TEMP\Flash_x86"
    
#Uninstalls any version of Java currently installed  
   
  #  $App = Get-WmiObject Win32_Product | Where {$_.Name -like "*Flash*"}
  #  $App.Uninstall() 
    

#Identifies whether or not it is a x86/x64 OS
    
    $OSInfo = Get-Wmiobject Win32_OperatingSystem -ErrorAction SilentlyContinue

                
        If ($OSInfo.OSArchitecture -eq "64-Bit")
        {
            
            If (!(Test-Path -path $destinationFolder_x64))
            {
                        
                New-Item $destinationFolder_x64 -Type Directory -Force
            }

                Copy-Item -Path $sourcefile_x64 -Destination $destinationFolder_x64 

                Invoke-Command  -ScriptBlock {cmd /c "start /wait msiexec.exe /i C:\TEMP\Flash_x64\install_plugin.msi /qn"}
                Invoke-Command  -ScriptBlock {cmd /c "start /wait msiexec.exe /i C:\TEMP\Flash_x64\install_active_x.msi /qn"}       
        }

 <#If x86 OS, creates one directory and copies all files for only x86 Java installation      
       
        If ($OSInfo.OSArchitecture -eq "32-Bit")
        {
            If (!(Test-Path -path $destinationFolder_x86))
            {
                New-Item $destinationFolder_x86 -Type Directory -Force
            }

            Copy-Item -Path $sourcefile_x86 -Destination $destinationFolder_x86
            Invoke-Command -ScriptBlock {msiexec.exe /i C:\TEMP\Java_x86\jre1.8.9_71.msi /qn}
        }
        Write-Host "Completed"
    
#>