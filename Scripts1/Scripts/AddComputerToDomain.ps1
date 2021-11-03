$domain = "AREA52.AFNOAPPS.USAF.MIL"
$password = 'Whateveryoumakeit' | ConvertTo-SecureString -asPlainText -Force
$username = 'Whateveritis' 
$credential = New-Object System.Management.Automation.PSCredential($username,$password)
$cred = get-credential
$systemName = $env:COMPUTERNAME
#if ($systemName -like "XLWU*") {
    Invoke-Command -Computername xlwul-461pxt -scriptblock {Add-Computer -DomainName $domain -Credential $cred -OUPath "OU=Tyndall AFB Computers,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" -Force -ErrorVariable addresult}
    if ($addresult -ne $null) {
        THROW 'ERROR'
        }
    
Else {
    THROW 'ERROR'
    }