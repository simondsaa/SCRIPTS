$domain = "AREA52.AFNOAPPS.USAF.MIL"
$password = 'Whateveryoumakeit' | ConvertTo-SecureString -asPlainText -Force
$username = 'Whateveritis' 
$credential = New-Object System.Management.Automation.PSCredential($username,$password)
$systemName = $env:COMPUTERNAME
if ($systemName -like "RKMF*") {
    Add-Computer -DomainName $domain -Credential $credential -OUPath "OU=Nellis AFB Computers,OU=Nellis AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" -Force -ErrorVariable addresult
    if ($addresult -ne $null) {
        THROW 'ERROR'
        }
    }
Else {
    THROW 'ERROR'
    }