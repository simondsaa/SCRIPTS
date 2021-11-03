#######################################################
##
## WakeUp-Machines.ps1, v1.1, 2012
##
## Created by Matthijs ten Seldam, Microsoft
##
#######################################################
 
<#
.SYNOPSIS
Starts a list of physical machines by using Wake On LAN.
 
.DESCRIPTION
WakeUp-Machines starts a list of servers using Wake On LAN magic packets. It then sends echo requests to verify that the machine has TCP/IP connectivity. It waits for a specified amount of echo replies before starting the next machine in the list.
 
.PARAMETER Machines
The name of the file containing the machines to wake.
 
.PARAMETER TimeOut
The number of seconds to wait for an echo reply before continuing with the next machine.
 
.PARAMETER Repeat
The number of echo requests to send before continuing with the next machine.
 
.EXAMPLE
WakeUp-Machines machines.csv
 
.EXAMPLE
WakeUp-Machines c:\tools\machines.csv
 
.INPUTS
None
 
.OUTPUTS
None
 
.NOTES
Make sure the MAC addresses supplied don't contain "-" or ".".
 
The CSV file with machines must be outlined using Name, MAC Address and IP Address with the first line being Name,MacAddress,IpAddress.
See below for an example of a properly formatted CSV file.
 
Name,MacAddress,IpAddress
Host1,A0DEF169BE02,192.168.0.11
Host3,AC1708486CA2,192.168.0.12
Host2,FDDEF15D5401,192.168.0.13
 
.LINK
http://blogs.technet.com/matthts
#>
 
param(
    [Parameter(Mandatory=$true, HelpMessage="Path to the CSV file containing the machines to wake.")]
    [string] $File,
    [Parameter(Mandatory=$false, HelpMessage="Number of unsuccesful echo requests before continuing.")]
    [int] $TimeOut=30,
    [Parameter(Mandatory=$false, HelpMessage="Number of successful echo requests before continuing.")]
    [int] $Repeat=10,
    [Parameter(Mandatory=$false, HelpMessage="Number of magic packets to send to the broadcast address.")]
    [int] $Packets=2
    )
 
#Set-StrictMode -Version Latest
 
#clear;Write-Host
 
## Read CSV file with machine names
#try
#{
    $Machines=Import-Csv $File
#}
#Catch
#{
#    Write-Host "$File file not found!";Write-Host
    #exit
#}
 
function Send-Packet([string]$MacAddress, [int]$Packets)
{
    <#
    .SYNOPSIS
    Sends a number of magic packets using UDP broadcast.
 
    .DESCRIPTION
    Send-Packet sends a specified number of magic packets to a MAC address in order to wake up the machine.  
 
    .PARAMETER MacAddress
    The MAC address of the machine to wake up.
 
    .PARAMETER
    The number of packets to send.
    #>
 
    try
    {
        $Broadcast = ([System.Net.IPAddress]::Broadcast)
 
        ## Create UDP client instance
        $UdpClient = New-Object Net.Sockets.UdpClient
 
        ## Create IP endpoints for each port
        $IPEndPoint1 = New-Object Net.IPEndPoint $Broadcast, 0
        $IPEndPoint2 = New-Object Net.IPEndPoint $Broadcast, 7
        $IPEndPoint3 = New-Object Net.IPEndPoint $Broadcast, 9
 
        ## Construct physical address instance for the MAC address of the machine (string to byte array)
        $MAC = [Net.NetworkInformation.PhysicalAddress]::Parse($MacAddress)
 
        ## Construct the Magic Packet frame
        $Frame = [byte[]]@(255,255,255, 255,255,255);
        $Frame += ($MAC.GetAddressBytes()*16)
 
        ## Broadcast UDP packets to the IP endpoints of the machine
        for($i = 0; $i -lt $Packets; $i++)
        {
            $UdpClient.Send($Frame, $Frame.Length, $IPEndPoint1) | Out-Null
            $UdpClient.Send($Frame, $Frame.Length, $IPEndPoint2) | Out-Null
            $UdpClient.Send($Frame, $Frame.Length, $IPEndPoint3) | Out-Null
            sleep 1;
        }
    }
    catch
    {
        $Error | Write-Error;
    }
}
 
$i=1
foreach($Machine in $Machines)
{
    $Name=$Machine.Name
    $MacAddress=$Machine.MacAddress
    $IPAddress=$Machine.IpAddress
 
    ## Send magic packet to wake machine
    Write-Progress -ID 1 -Activity "Waking up machine $Name" -PercentComplete ($i*100/$file.Count)
    Send-Packet $MacAddress $Packets
 
    $j=1
    ## Go into loop until machine replies to echo
    $Ping = New-Object System.Net.NetworkInformation.Ping
    do
    {
        $Echo = $Ping.Send($IPAddress)
        Write-Progress -ID 2 -ParentID 1 -Activity "Waiting for $Name to respond to echo" -PercentComplete ($j*100/$TimeOut)
        sleep 1
        
        if ($j -eq $TimeOut)
        {
            Write-Host "Time out expired, aborting.";Write-Host
            exit
        }
        $j++
    }
    while ($Echo.Status.ToString() -ne "Success")
 
    ## Machine is alive, keep sending for $Replies amount
    for ($k = 1; $k -le $Repeat; $k++)
    {
       Write-Progress -ID 2 -ParentID 1 -Activity "Waiting for $Name to respond to echo" -PercentComplete (100)
       Write-Progress -Id 3 -ParentId 2 -Activity "Receiving echo reply"  -PercentComplete ($k*100/$Repeat)
       sleep 1
    }
    $i++
    Write-Progress -Id 3 -Completed $true
    $Ping=$null
}