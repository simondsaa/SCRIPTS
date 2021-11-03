$computers=Get-Content -Path C:\Temp\g2.txt
foreach ($computer in $computers) {
    write-host ""
    write-host "Working on $Computer" -Foreground green     
    $Check_Connection = test-connection -computername $Computer -quiet -count 1
    If($Check_Connection -eq $true)
        {
        Write-Host "$Computer is on the network. Attempting to reset the BIOS password."
        (Get-WmiObject -computername $Computer -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface).SetBIOSSetting('Setup Password','<utf-16/>','<utf-16/>password')
        }
    Else
        {
            write-host "$Computer is not on the network." -Foreground yellow                                  
        } 
}
