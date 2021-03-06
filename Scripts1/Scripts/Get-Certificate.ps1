$Cert = Get-Credential

Add-Type -AssemblyName System.Security
    # You can do more filtering here if there are other cert requirements...
    #$ValidCerts = [System.Security.Cryptography.X509Certificates.X509Certificate2[]](dir Cert:\CurrentUser\My | where { $_.NotAfter -gt (Get-Date) })

    $ValidCerts | select FriendlyName, Thumbprint, subject, DnsNameList, Issuer, EnhancedKeyUsageList, NotAfter | fl
    
    # You could check $ValidCerts, and not do this prompt if it only contains 1...
    $Cert = [System.Security.Cryptography.X509Certificates.X509Certificate2UI]::SelectFromCollection(
    $ValidCerts,
    'Choose a certificate',
    'Choose a certificate',
    'SingleSelection'
    ) | select -First 1