
  
    #$computers = Get-Content \\xlwu-fs-05pv\Tyndall_PUBLIC\Patching\Java\JavaPatch.txt    
    $sourcefile_x64 = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Applications\Java\Java_8.91\x64\*"
    $sourcefile_x86 = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Applications\Java\Java_8.91\x86\*"
    $destinationFolder_x64 = "C:\TEMP\Java_x64"
    $destinationFolder_x86 = "C:\TEMP\Java_x86"
    
#Uninstalls any version of Java currently installed  
   
    $App = Get-WmiObject Win32_Product | Where {$_.Name -like "*Java*"}
    $App.Uninstall() 
    
#Identifies whether or not it is a x86/x64 OS
    
    $OSInfo = Get-Wmiobject Win32_OperatingSystem -ErrorAction SilentlyContinue

        
#If x64 OS, creates two directories and copies all files for both x64 and x86 Java installs
        
        If ($OSInfo.OSArchitecture -eq "64-Bit")
        {
            Write-Host "System is x64 bit. This will install both x86 and x64 versions of Java..."
            
            If (!(Test-Path -path $destinationFolder_x64))
            {
                        
            Write-Host "Creating TEMP\Java_x64 Directory"
                New-Item $destinationFolder_x64 -Type Directory -Force
            }

            If (!(Test-Path -path $destinationFolder_x86))
            {
            
            Write-Host "Creating TEMP\Java_x86 Directory"
                New-Item $destinationFolder_x86 -Type Directory -Force
            }

            Write-Host "Copying Files"

                Copy-Item -Path $sourcefile_x64 -Destination $destinationFolder_x64 
                Copy-Item -Path $sourcefile_x86 -Destination $destinationFolder_x86

            Write-Host "Copy complete. Beginning installation of Java x64..."

                Invoke-Command  -ScriptBlock {cmd /c "start /wait msiexec.exe /i C:\TEMP\Java_x64\jre1.8.0_71.msi /qn"}

            Write-Host "Java x64 installation complete. Beginning installation of Java x86"

                Invoke-Command  -ScriptBlock {cmd /c "start /wait msiexec.exe /i C:\TEMP\Java_x86\jre1.8.0_71.msi /qn"}       
        }

 #If x86 OS, creates one directory and copies all files for only x86 Java installation      
       
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
    
