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

$colProplist = "sn","givenname"

foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}

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

#Enable WMI Remoting
$WinRMTest =  Get-Service winrm


If ($WinRMTest.Status -eq "Running") {
$WinRM = "Active"
}
ELSE {
C:\windows\system32\winrm.cmd quickconfig -quiet
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
$OutputObj | add-member -MemberType NoteProperty -Name EDIPI -Value $env:USERNAME
$OutputObj | add-member -MemberType NoteProperty -Name FIRST.LAST -Value $objName
$OutputObj | add-member -MemberType NoteProperty -Name DHCP -Value $objNet.dhcpserver
$OutputObj | add-member -MemberType NoteProperty -Name MAC -Value $objNet.macaddress
$OutputObj | Add-Member -membertype NoteProperty -name Location -value $objComputer.location[0]
$OutputObj | Add-Member -MemberType NoteProperty -Name Manufacturer -Value (Get-WmiObject -class win32_bios).Manufacturer
$OutputObj | Add-Member -MemberType NoteProperty -Name SerialNumber -Value (Get-WmiObject -class win32_bios).SerialNumber
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


#Current Computer Stats
$Outarray | select Date,ComputerName,First.Last,EDIPI,OS_Version,OS_Architecture,Server,Netlogon,Remoting,IPAddress,MAC,DefaultIPGateway,DHCP,Manufacturer,SerialNumber,Location,DayOfYear | Export-csv "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\WW2 Recovery\$env:computername.csv" -NoTypeInformation




# Append pre-existing Current file to Archive file
# Overwrite pre-existing Current file

