$Comp = (Get-Content C:\Users\timothy.brady\Desktop\Comps.txt)
Get-WmiObject -Class Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $Comp | Select IPAddress, MACAddress