$Comp = Read-Host
Invoke-Command -cn $Comp {CD "HKLM:\System\CurrentControlSet\Control\Terminal Server" | Set-ItemProperty -name AllowRemoteRPC -value 1}
CD C:\Windows\System32
Query Session /Server:$Comp