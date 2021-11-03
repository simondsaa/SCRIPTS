#####################################################################################################
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#####################################################################################################
#####################################################################################################
# Script: Stats.ps1
# Purpose: Gathers and creates files in .CSV format to be combined into a single object containing all
#          of the important information on Computers, Users, and ULUC2 Machines. Replaced .TXT stat
#          files being gathered by batch files, as these cannot be combined manipulated within PoSH.
# Creator: SSgt Crill, Christian 325 CS/SCOO
# Date: 7/29/2015
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
cls
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#####################################################################################################
#                                   ACTIVE DIRECTORY SCAN CODEBLOCK                                 #
#         Prefer to combine these into one scan/object instead of two seperate scans/objects        #
#####################################################################################################
#####################################################################################################

#Current logon name
$objWho = [environment]::username
#ADUC INFO PULL
$ADsPath = "LDAP://OU=Tyndall AFB Users,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$ADObj = New-Object PSObject;
$ADInfo = @();
$objEDIPI =   [environment]::username
$strFilter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$objWho))"


$ADObj = New-Object PSObject;
$objDomain = New-Object System.DirectoryServices.DirectoryEntry
$objOU = New-Object System.DirectoryServices.DirectoryEntry($ADsPath)

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objOU
$objSearcher.PageSize = 1000
$objSearcher.Filter = $strFilter
$objSearcher.SearchScope = "Subtree"

#$colProplist = "sn","givenname"

#foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}

$CollSystems = $objSearcher.FindAll()

#Configure $ADObj to hold all AD data 
 ForEach($objResult in $CollSystems) 
    {
        $objItem = $objResult.Properties; $AllComputerNames +=$objItem.name;$ADObj = New-Object PSObject;$ADObj | 
        Add-Member NoteProperty SurName $objItem.sn -Force; $ADObj | 
        Add-Member NoteProperty givenname $objItem.givenname -Force; $ADObj
            }
#Username variable combination
$ObjNameArray =  ($ADObj.givenname)+($ADObj.Surname)
#Combine onto one line with a . between
$objName = $ObjNameArray -join '.'

# Computer info
# ADUC Info Grab
$domain = "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$objDomain = [adsi]("LDAP://" + $domain)
$Computer = $env:COMPUTERNAME

$search = New-Object System.DirectoryServices.DirectorySearcher
$search.SearchRoot = $objDomain
$search.Filter = "(&(objectClass=computer)(cn=$Computer))"
$search.SearchScope = "Subtree"
$results = $search.FindAll()

    ForEach($item in $results){
        $objComputer = $item.GetDirectoryEntry()
            }

cls
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
##########################################################################################################
#                                           BODY                                                    
# Gathers all information into a single object ($OutArray)
# All statistic files export information from this single object
# Known issue wih IPAdress / DefaultIPGateWay / Location where they are displayed in CSV as System.Objects
# This is resolved by pumping them into $OutputObj then adding them to $OutputArray
# ULUC2 machines are called at the end of the script for their popup information and their stat files
###########################################################################################################
###########################################################################################################
# Create blank line
$nl = [Environment]::newline

#WMI Remoting Test
$WinRMTest =  Get-Service winrm


If ($WinRMTest.Status -eq "Running") {
$WinRM = "Active"
}
ELSE {
$WinRMTest
         IF ($WinRMTest.Status -eq "Running") {
         $WinRM = "Active"
         }
         ELSE {
         $WINRM = "Inactive"
         }
}
$NetlogonTest = Get-Service Netlogon

If ($NetLogonTest.Status -eq "Running") {
$Netlogon = "Active"
}
ELSE {
$Netlogon = "Inactive"
         }


$ServerTest = Get-Service Server

If ($NetLogonTest.Status -eq "Running") {
$Server = "Active"
}
ELSE {
$Server = "Inactive"
         }



#Creates filename for csv FIRST.LAST_EDIPI
$arrayUserName = "$objName","$env:username"
$separator = "_"

$UserInfo_FileName = [string]::Join($separator,$arrayUserName)

#Grabs network information
$networkinfo = (Get-WMIObject -Class Win32_NetworkAdapterConfiguration -filter "IPEnabled = true")

#Converts objects with collections into a single string to be displayed by a CSV

$Outarray = @()
foreach($objNet in $networkinfo) {

    $groups = $objitem.memberof
    ForEach ($group in $groups)
    {
        $Trim = $group.TrimStart("CN=")
        $final += $Trim.Split(",")[0] + ", "
    }

$OutputObj = New-Object -TypeName PSobject
$OutputObj | Add-Member -MemberType NoteProperty -Name IPAddress -Value $objNet.IPAddress[0]
$OutputObj | Add-Member -MemberType NoteProperty -Name DefaultIPGateway -Value $objNet.DefaultIPGateway[0]
$OutputObj | add-member -MemberType NoteProperty -Name Date -Value (get-date)
$OutputObj | add-member -MemberType NoteProperty -Name Remoting -Value $WinRM
$OutputObj | add-member -MemberType NoteProperty -Name Server -Value $Server
$OutputObj | add-member -MemberType NoteProperty -Name NetLogon -Value $NetLogon
$OutputObj | add-member -MemberType NoteProperty -Name DayOfYear -Value (get-date).DayOfYear
$OutputObj | add-member -MemberType NoteProperty -Name ComputerName -Value $env:COMPUTERNAME 
$OutputObj | add-member -MemberType NoteProperty -Name OS_Version -Value (get-itemproperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ProductName).productname
$OutputObj | add-member -MemberType NoteProperty -Name OS_Architecture -Value (Get-WmiObject Win32_OperatingSystem).OSArchitecture
$OutputObj | add-member -MemberType NoteProperty -Name Organization -Value $objitem.o[0]
$OutputObj | add-member -MemberType NoteProperty -Name PhoneNumber -Value $objitem.telephonenumber[0]
$OutputObj | add-member -MemberType NoteProperty -Name EDIPI -Value $env:USERNAME
$OutputObj | add-member -MemberType NoteProperty -Name FIRST.LAST -Value $objName
$OutputObj | add-member -MemberType NoteProperty -Name DHCP -Value $objNet.dhcpserver
$OutputObj | add-member -MemberType NoteProperty -Name MAC -Value $objNet.macaddress
$OutputObj | Add-Member -membertype NoteProperty -name Location -value $objComputer.location[0]
$OutputObj | Add-Member -MemberType NoteProperty -Name Manufacturer -Value (Get-WmiObject -class win32_bios).Manufacturer
$OutputObj | Add-Member -MemberType NoteProperty -Name SerialNumber -Value (Get-WmiObject -class win32_bios).SerialNumber
$OutputObj | add-member -MemberType NoteProperty -Name Groups -Value $final
$Outarray += $OutputObj
}
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#####################################################################################################
#                                          File Outputs                                             #
#####################################################################################################
#####################################################################################################

# Check Archive file
# If exceeds 10kb delete Archive File

$Comp_csv = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Archives\Computer_Stats\Concurrent\$env:computername.csv" 
$User_csv = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Archives\User_Stats\Concurrent\$UserInfo_FileName.csv" 

$Archive_Comp = Import-csv $Comp_csv -ErrorAction SilentlyContinue
$Archive_User = Import-Csv $User_csv -ErrorAction SilentlyContinue

#Computer stats
IF (Test-path $Comp_csv) {
         IF( (get-item $Comp_csv).Length -gt 10240) {
            $a = $Outarray 
            $a | select Date,ComputerName,First.Last,Organization,EDIPI,OS_Version,OS_Architecture,Server,Netlogon,Remoting,IPAddress,MAC,DefaultIPGateway,DHCP,Manufacturer,SerialNumber,Location,DayOfYear | export-csv "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Archives\Computer_Stats\Concurrent\$env:computername.csv" -NoTypeInformation
            $a = $null
    } ELSE {
            $a = $Outarray + $Archive_Comp
            $a | select Date,ComputerName,First.Last,Organization,EDIPI,OS_Version,OS_Architecture,Server, Netlogon,Remoting,IPAddress,MAC,DefaultIPGateway,DHCP,Manufacturer,SerialNumber,Location,DayOfYear | export-csv "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Archives\Computer_Stats\Concurrent\$env:computername.csv" -NoTypeInformation
            $a = $null
            }
} ELSE {
$a = $Outarray + $Archive_Comp
$a | select Date,ComputerName,First.Last,Organization,EDIPI,OS_Version,OS_Architecture,Server,Netlogon,Remoting,IPAddress,MAC,DefaultIPGateway,DHCP,Manufacturer,SerialNumber,Location,DayOfYear |  export-csv "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Archives\Computer_Stats\Concurrent\$env:computername.csv" -NoTypeInformation
$a = $null
}

#User Stats
IF (Test-path $User_csv) {
         IF( (get-item $User_csv).Length -gt 10240) {
            $a = $Outarray
            $a | select Date,ComputerName,Organization,FIRST.LAST,EDIPI,PhoneNumber,Groups,IPAddress,MAC,Manufacturer,SerialNumber,Location,DayOfYear| export-csv "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Archives\User_Stats\Concurrent\$UserInfo_FileName.csv" -NoTypeInformation
            $a = $null
     } ELSE {
            $a = $Outarray + $Archive_User
            $a | select  Date,ComputerName,Organization,FIRST.LAST,EDIPI,PhoneNumber,Groups,IPAddress,MAC,Manufacturer,SerialNumber,Location,DayOfYear | export-csv "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Archives\User_Stats\Concurrent\$UserInfo_FileName.csv" -NoTypeInformation
            $a = $null
             }
}
ELSE {
$a = $Outarray + $Archive_User
$a | select   Date,ComputerName,Organization,FIRST.LAST,EDIPI,PhoneNumber,Groups,IPAddress,MAC,Manufacturer,SerialNumber,Location,DayOfYear | export-csv "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Archives\User_Stats\Concurrent\$UserInfo_FileName.csv" -NoTypeInformation
$a = $null
}




#Current Computer Stats
$Outarray | select Date,ComputerName,First.Last,Organization,EDIPI,OS_Version,OS_Architecture,Server,Netlogon,Remoting,IPAddress,MAC,DefaultIPGateway,DHCP,Manufacturer,SerialNumber,Location,DayOfYear | Export-csv "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Current\Computer_Stats\Windows\$env:computername.csv" -NoTypeInformation
#Computer Print to Temp

$Outarray | select Date,ComputerName,First.Last,Organization,EDIPI,OS_Version,OS_Architecture,Server,Netlogon,Remoting,IPAddress,MAC,DefaultIPGateway,DHCP,Manufacturer,SerialNumber,Location,DayOfYear | Export-csv "C:\Temp\System_Info.csv" -NoTypeInformation -force

#Current User Stats
$Outarray | select Date,ComputerName,Organization,FIRST.LAST,EDIPI,PhoneNumber,Groups,IPAddress,MAC,Manufacturer,SerialNumber,Location,DayOfYear | export-csv "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Current\User_Stats\$UserInfo_FileName.csv" -NoTypeInformation




# Append pre-existing Current file to Archive file
# Overwrite pre-existing Current file

#User stats



# Script is done if this is not a UL Client
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                                                                                   #
#                                    Tests if UL is installed on machine
#####################################################################################################
#                                     Unit Level Unit Command and Control                           
#                                                  ULUC2                          
# Popup displays information to the user, and creates a stats file for UL clients                  
#####################################################################################################

$x86 = test-path "C:\Program Files\IIMS"
$x64 = test-path "C:\Program Files (x86)\IIMS"
IF ($x86 -or $x64) {


##########################
#  Set Window Parameters #
##########################

#Adjusts windows size for UL Users
$pshost = get-host
$pswindow = $pshost.ui.rawui

$newsize = $pswindow.buffersize
$newsize.height = 51
$newsize.width = 80
$pswindow.buffersize = $newsize

$newsize = $pswindow.windowsize
$newsize.height = 50
$newsize.width = 80
$pswindow.windowsize = $newsize

$pswindow.windowtitle = "UL System Info"

############################
#  User Display Parameters #
############################

#Displays below information for UL Clients

# Java_home locations for x64 and x86
$j86 = "C:\Program Files\IIMS"
$j64 = "C:\Program Files (x86)\IIMS"

#Windows x86 or x64
$WinVer = (Get-WmiObject Win32_OperatingSystem).OSArchitecture

#
#
# Objects below here are displayed to the users in text form
#
# 
	write-host $winVer "ULUC2 Client" -Foregroundcolor Yellow
	write-host "Client Version: $env:ul_version" -Foregroundcolor Yellow
	write-host "Java Location: $env:Java_home" -foregroundcolor Yellow
$nl
	write-host "	  	 Tyndall AFB		" -foregroundcolor Green
	Write-host "   NIPR Unit Level/Command and Control		" -foregroundcolor Green
	Write-host "	   HTTPS://52XLWU-WS-001		" -FOREGROUNDCOLOR Yellow
$nl
$nl
	write-host "For any system issues between 0730 - 1630 please contact 283-8230" -Foregroundcolor Green
	write-host "For any client issues between 0730 - 1630 please contact 283-4896" -Foregroundcolor Green
	write-host "For any any issues outside of this range please contact 283-4896" -Foregroundcolor Green

$nl
$nl
	write-host "::ATTENTION::  $nl If you plan to replace your computer (tech refresh), $nl the new computer will need ULUC2 installed" -Foregroundcolor cyan -BackgroundColor darkcyan
$nl
	write-host "::ATTENTION::  $nl For Alerter or IIMS Java errors please restart your machine prior to calling $nl in an issue, if it is not resolved please proceed." -backgroundcolor darkRed -foregroundcolor yellow
$nl

$nl
	write-host "::ATTENTION::  $nl The map is being worked on and may display several warning indicators, please ignore these warnings as the map is adjusted." -backgroundcolor darkRed -foregroundcolor yellow

#
#
# End Objects being displayed to users
#
#


############################
#  Create UL stat objects  #
############################

# Adding UL objects to OutputArray
$Outarray | Add-Member -MemberType NoteProperty -Name Java_Home -Value $env:Java_home
$Outarray | Add-Member -MemberType NoteProperty -Name UL_Version -Value $env:ul_version


#Creates filename for csv 64/32-bit_COMPUTERNAME.csv
$arrayFileName = "$winver","$env:computername"
$separator = "_"

$UL_FileName = [string]::Join($separator,$arrayFileName)

#UL Stats
$Outarray | select Date,FIRST.LAST,Organization,PhoneNumber,EDIPI,UL_Version,Java_Home,OS_Architecture,Computername,IpAddress,Remoting,MAC,Location,DayOfYear | export-csv "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Current\UL_Clients\$UL_FileName.csv" -NoTypeInformation 


$nl
$nl
$nl

# Forces window open until user recognition 
Write-Host "Press any key to continue ..."

$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

}

ELSE {
exit
}

