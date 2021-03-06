$Computers = Get-Content C:\work\bowling.txt 
ForEach ($Computer in $Computers){If(Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {$NIC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $Computer -ErrorAction SilentlyContinue
        write-Host "$Computer is Pingable","|IP:"$NIC.IPAddress,"|MAC:"$NIC.MACAddress}
Else    {write-host -ForegroundColor Red "$Computer is Not Pingable"}}