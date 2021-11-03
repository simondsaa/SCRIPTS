$Computers = Get-Content C:\Users\timothy.brady\Desktop\Comps.txt 
ForEach ($Computer in $Computers){If(Test-Connection $Computer -quiet -count 1)
        {$NIC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $Computer -ErrorAction SilentlyContinue
        write-Host "$Computer is Pingable","|IP:"$NIC.IPAddress,"|MAC:"$NIC.MACAddress}
Else    {write-host -ForegroundColor Red "$Computer is Not Pingable"}}