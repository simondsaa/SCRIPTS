#$Comps = Get-Content C:\Users\1180219788.adm\Desktop\MDG_Computers.txt
       
ForEach ($Comp in $Comps)
    {       
        If (Test-Connection $Comp -quiet -BufferSize 64 -Ea 0 -count 1)
            { 
                netsh advfirewall firewall add rule name="ICMP Allow incoming V4 echo request" protocol="icmpv4:8,any" dir=in action=allow
                $result = "$Comp is accessible.  Enabling ICMP Echo."
                $result | Out-File -Verbose -FilePath C:\Users\1180219788.adm\Desktop\ICMP_Success.txt -Append -Force
            } 
            Else
                {
                    $result = "$Comp is not accessible."
                    $result | Out-File -Verbose -filepath C:\Users\1180219788.adm\Desktop\ICMP_failed.txt -Append -Force
                }  
    }    