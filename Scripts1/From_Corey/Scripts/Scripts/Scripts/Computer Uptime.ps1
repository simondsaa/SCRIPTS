$Computers = Get-Content "C:\Users\timothy.brady\Desktop\Comps.txt"
ForEach ($Computer in $Computers){
    $lastboottime = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).LastBootUpTime
    $sysuptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($lastboottime) 
    Write-Host "$computer has been up for: " $sysuptime.days "days" $sysuptime.hours "hours" $sysuptime.minutes "minutes" $sysuptime.seconds "seconds"
    }