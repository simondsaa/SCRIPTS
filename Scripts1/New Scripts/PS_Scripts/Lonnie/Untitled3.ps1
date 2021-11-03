        #$MAC = Get-WmiObject Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $CompName | select MACAddress -ErrorAction SilentlyContinue
        $IPMAC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $CompName
        $MAC = ($IPMAC | where { $_.IpAddress -eq $IPAddress}).MACAddress