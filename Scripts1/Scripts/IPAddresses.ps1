$servers = get-content C:\Users\1252862141.adm\Desktop\Scripts1\IPAddresses.txt
foreach ($Server in $Servers)
{
    $Addresses = $null
    try {
        #$Addresses = [System.Net.Dns]::GetHostAddresses("$Server").IPAddressToString
        $Addresses = nslookup $Server
    }
    catch { 
        $Addresses = ""
    }
    foreach($Address in $Addresses) {
        Write-Output $Server, $Address >> C:\Users\1252862141.adm\Desktop\IPAddresses.txt
    }
}