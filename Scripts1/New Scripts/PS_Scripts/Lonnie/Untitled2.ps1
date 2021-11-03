function Get-IPMAC
{
<#
        .Synopsis
        Function to retrieve IP & MAC Address of a Machine.
        .DESCRIPTION
        This Function will retrieve IP & MAC Address of local and remote machines.
        .EXAMPLE
        PS>Get-ipmac -ComputerName viveklap
        Getting IP And Mac details:
        --------------------------

        Machine Name : TESTPC
        IP Address : 192.168.1.103
        MAC Address: 48:D2:24:9F:8F:92
        .INPUTS
        System.String[]
        .NOTES
        Author - Vivek RR
        Adapted logic from the below blog post
        "http://blogs.technet.com/b/heyscriptingguy/archive/2009/02/26/how-do-i-query-and-retrieve-dns-information.aspx"
#>

Param
(
    #Specify the Device names
    [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            Position=0)]
    [string[]]$ComputerName
)
    Write-Host "Getting IP And Mac details:`n--------------------------`n"
    foreach ($Inputmachine in $ComputerName )
    {
        if (!(test-Connection -Cn $Inputmachine -quiet))
            {
            Write-Host "$Inputmachine : Is offline`n" -BackgroundColor Red
            }
        else
            {

            $MACAddress = "N/A"
            $IPAddress = "N/A"
            $IPAddress = ([System.Net.Dns]::GetHostByName($Inputmachine).AddressList[0]).IpAddressToString
            #$IPMAC | select MACAddress
            $IPMAC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Inputmachine
            $MACAddress = ($IPMAC | where { $_.IpAddress -eq $IPAddress}).MACAddress
            Write-Host "Machine Name : $Inputmachine`nIP Address : $IPAddress`nMAC Address: $MACAddress`n"
      
            }
    }
}