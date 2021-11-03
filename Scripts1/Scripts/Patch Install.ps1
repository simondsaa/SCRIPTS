<#-----------------------------------------------------------------------------------------#
 #                                  Written by SrA David Roberson                          #
 #                                  325 Communications Squadron                            #
 #                                  Tyndall AFB, Panama City, FL                           #
 #                                     Created 17 July 2014                                #
 #-----------------------------------------------------------------------------------------#>

# Get the Operating System architecture
$OS = (Get-WmiObject Win32_OperatingSystem).OSArchitecture

# Specify the location of the *.msu files
$updatedir = "\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC Admin\17-10"

$files = Get-ChildItem $updatedir -Recurse

If ($OS -eq "64-bit")
    {
    foreach ($file in $files)
        {
        If ($file.Name -like "*x64*")
            {
                Write-Host -ForegroundColor Gray "Installing update $file ..."
                $fullname = $file.fullname
                # Specify the command line parameters for wusa.exe
                $parameters = $fullname + " /quiet /norestart"
                # Start wusa.exe and pass in the parameters
                $install = [System.Diagnostics.Process]::Start( "wusa",$parameters )
                $install.WaitForExit()
                Write-Host -ForegroundColor Green "Finished installing $file"
            }
        }
    }
ElseIf ($OS -eq "32-bit")
    {
    foreach ($file in $files)
        {
        If ($file.Name -like "*x86*")
            {
                Write-Host -ForegroundColor Gray "Installing update $file ..."
                $fullname = $file.fullname
                # Specify the command line parameters for wusa.exe
                $parameters = $fullname + " /quiet /norestart"
                # Start wusa.exe and pass in the parameters
                $install = [System.Diagnostics.Process]::Start( "wusa",$parameters )
                $install.WaitForExit()
                Write-Host -ForegroundColor Green "Finished installing $file"
            }
        }
    }
Else
    {
        Write-Host -ForegroundColor Red "OS Architecture UNKNOWN, quitting..."
        Write-Host -ForegroundColor Yellow "Press Any Key to Exit..."
        $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp") > $null
        EXIT
    }
Write-Host "Installation Complete"
#Write-Host "System will reboot in 60 seconds..."
#Shutdown -r -f -t 60