$Comp = XLWUW-491s6m
Get-WmiObject -Class Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $Comp | Select IPAddress, MACAddress