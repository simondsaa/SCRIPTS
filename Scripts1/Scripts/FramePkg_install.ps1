$Comps = "xlwuw-421nkc"
#$Comps = Get-Content C:\Users\1180219788A\Desktop\rogue.txt

ForEach($Comp in $Comps)
    {
    $Path = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Applications\McAfee_Install\ePO9FramePkg\5.4FramePkg1"
    $Executable = "FramePkg.exe"
    $Switches = "/INSTALL=AGENT /S /FORCEINSTALL"

    $Cmd = "\\$Comp\c$\windows\temp\$Executable $Switches"
    $FullPath = $Path + "\" + $Executable
    $ErrorActionPreference="SilentlyContinue"

    $Ping = New-Object System.Net.NetworkInformation.Ping
    $Reply = $Ping.send($Comp)
        If ($Reply.status -eq "Success")
            {
            #Write-Host "$Comp - System online" -ForegroundColor GREEN
            Copy-Item $FullPath -Destination \\$Comp\c$\windows\temp
    
            Trap { Write-Warning "There was an error connecting to the remote Computer or creating the process"; continue }      
    
            Write-Host "Connecting to $Comp"
    
            $Wmi=([wmiclass]"\\$Comp\root\cimv2:win32_process")  
            #bail out if the object didn't get created 
    
            If (!$Wmi) { return }  
    
            $Remote=$Wmi.Create($Cmd)  
    
            If ($Remote.returnvalue -eq 0) 
                { Write-Host "Successfully launched on $Comp" -ForegroundColor GREEN } 
            Else { Write-Host "Failed to launch $Cmd on $Comp. ReturnValue is" $Remote.ReturnValue -ForegroundColor RED } }
            Else { Write-Host "$Comp - System offline" -ForegroundColor RED } }
