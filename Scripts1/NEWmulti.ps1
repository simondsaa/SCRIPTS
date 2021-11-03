#===================================================================

Function SendMessage
{
    REG ADD "\\$Computer\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v AllowRemoteRPC /t REG_DWORD /d 1 /f
    $Message = Read-Host "Message"
    $SendMsg = MSG console /Server:$Computer /Time:6000 $Message
}

#===================================================================

Function LockComputer
{
    $Lock = TSDiscon Console /Server:$Computer
}

#===================================================================

Function LogoffUser
{
    $Logoff = Reset Session Console /Server:$Computer
}

#===================================================================

Function RebootComputer
{
    Write-Host " "
    Write-Host "1 - 1 Minute"
    Write-Host "2 - 10 Minutes"
    Write-Host "3 - 30 Minutes"
    Write-Host "4 - 1 Hour"
    Write-Host "5 - Now"
    $Ans = Read-Host "Make Selection"
    
    If ($Ans -eq 1)
    {
        Write-Host
        $Time = 60
    }
    If ($Ans -eq 2)
    {
        Write-Host
        $Time = 600
    }
    If ($Ans -eq 3)
    {
        Write-Host
        $Time = 1800
    }
    If ($Ans -eq 4)
    {
        Write-Host
        $Time = 36000
    }
    If ($Ans -eq 5)
    {
        Write-Host
        $Time = 0
    }
    If ($Time -gt 0)
    {
        $RTime = $Time/60
    }
    $Input = Read-Host "Comment"
    $Comment = $Input+": You will be rebooted in $RTime minute(s), please save all work"
    $Shutdown = Shutdown /r /f /m \\$Computer /t $Time /c $Comment
}

#===================================================================

Function EnableLocalAdmin
{
param($computer="localhost", $a, $user, $password, $help, $i, $c, $f, $work, $in)
function work() 
{
$EnableUser = 512

$DisableUser = 2

if(Test-Connection -ComputerName $c -Quiet)

  {

if(!$user)

      {

       $(Throw 'A value for $user is required.

       Try this: EnableDisableUser.ps1 -help ?')

        }

     

$ObjUser = [ADSI]"WinNT://$c/$user"
 
switch($a)

{

 "e" {

      $objUser.setpassword($password)

      $objUser.description = "Enabled Account"

      $objUser.userflags = $EnableUser

      $objUser.setinfo()

       }

 "d" {

      $objUser.description = "Disabled Account"

      $objUser.userflags = $DisableUser

      $objUser.setinfo()

       }

 DEFAULT

        {

             "You must supply a value for the action.

             Try this: EnableDisableUser.ps1 -help ?"

            }
}
}
Else
      {

       $(Write-Error 'Could not change local admin password.')

       }

}

function funHelp()

{

$helpText=@"

DESCRIPTION:

NAME: EnableDisableUser.ps1

Enables or Disables a local user on either a local or remote machine.

PARAMETERS:

-computer Specifies the name of the computer upon which to run the script

-a(ction) Action to perform < e(nable) d(isable) >

-user     Name of user to create

-help     prints help file

 

SYNTAX:

EnableDisableUser.ps1

Generates an error. You must supply a user name

 

EnableDisableUser.ps1 -computer MunichServer -user myUser

-password Passw0rd^&! -a e

 

Enables a local user called myUser on a computer named MunichServer

with a password of Passw0rd^&!

 

EnableDisableUser.ps1 -user myUser -a d

Disables a local user called myUser on the local machine

 

EnableDisableUser.ps1 -help ?

 

Displays the help topic for the script

 

"@

$helpText

exit

}

$a = Read-Host "Enter E to ENABLE or D to DISABLE"

if(!(($a -eq "e") -or ($a -eq "d")))

      {

       $(Throw 'Input value must be (e) for enable or (d) for disable')

       }

$user = "usaf_admin"

if($a -eq "e")

          {
                $password = "zaq1XSW@zaq1XSW@"
        
                if(!$password)

                {

                    $(Throw 'a value for $password is required.

                     Try this: EnableDisableUser.ps1 -help ?')

                }
           }

$in = Read-Host "Enter S for one PC or M for multiple"

if(!(($in -eq "s") -or ($in -eq "m")))

      {

       $(Throw 'Input value must be (s) for single computer or (m) for multiple computers')

       }


switch($in)

{

 "s" {
        $c = Read-Host "PC Name"
        work([string]$c)
       }

 "m" {
        $f = Read-Host "Enter Path"
        $FileExists = Test-Path $f 
        If ($FileExists -eq $True) 
                    { 
                        $i = Get-Content $f
                        foreach ($c in $i)
                        {$c + "`n" + "=========================="; work([string]$c)}
                    }
        Else
                    {

                        $(Write-Error 'Path to input file is not correct 

                          or is not accessable with the current user.')

                    }
       }
}

if($help){ "Obtaining help ..." ; funhelp }

function work() 
{
$EnableUser = 512

$DisableUser = 2

if(Test-Connection -ComputerName $c -Quiet)

  {

if(!$user)

      {

       $(Throw 'A value for $user is required.

       Try this: EnableDisableUser.ps1 -help ?')

        }

     

$ObjUser = [ADSI]"WinNT://$c/$user"
 
switch($a)

{

 "e" {

      $objUser.setpassword($password)

      $objUser.description = "Enabled Account"

      $objUser.userflags = $EnableUser

      $objUser.setinfo()

       }

 "d" {

      $objUser.description = "Disabled Account"

      $objUser.userflags = $DisableUser

      $objUser.setinfo()

       }

 DEFAULT

        {

             "You must supply a value for the action.

             Try this: EnableDisableUser.ps1 -help ?"

            }
}
}
Else
      {

       $(Write-Error 'Can not contact computer. 

       It is either currently offline or not reachable through the network.')

       }

}
}

#===================================================================

Function MissedCall
{
$Comp = Read-Host "User Name"
    If ($Comp -eq "Pelletier"){$Compname = "XLWUW-491S33"}
    ElseIf ($Comp -eq "Grainger"){$Compname = "XLWUW-491S8K"}
    ElseIf ($Comp -eq "Ballentine"){$Compname = "XLWUW-432LBH"}
    ElseIf ($Comp -eq "Foster"){$Compname = "XLWUW-491S64"}
    ElseIf ($Comp -eq "Mowry"){$Compname = "XLWUW-491S40"}
    ElseIf ($Comp -eq "Lozada"){$Compname = "XLWUW-491S7T"}
    ElseIf ($Comp -eq "Brown"){$Compname = "XLWUW-491S96"}
    ElseIf ($Comp -eq "Barnett"){$Compname = "XLWUW-491S8S"}
    ElseIf ($Comp -eq "Cain"){$Compname = "XLWUW-47168P"}
    ElseIf ($Comp -eq "Simonds"){$Compname = "XLWUW-491S35"}
    ElseIf ($Comp -eq "Ray"){$Compname = "XLWUW-471P8W"}
    ElseIf ($Comp -eq "Rick"){$Compname = "XLWUW-491S50"}
    ElseIf ($Comp -eq "Lewis"){$Compname = "XLWUW-4208TT"}
    ElseIf ($Comp -eq "Carnall"){$Compname = "XLWUW-471P8F"}
    

$User = Get-WmiObject Win32_ComputerSystem -Property Username -Comp $Compname
    If ($User.UserName -eq "AREA52\1383807847N"){$Name = "Pelletier"}
    ElseIf ($User.UserName -eq "AREA52\1253515879N"){$Name = "Grainger"}
    ElseIf ($User.UserName -eq "AREA52\1395576280N"){$Name = "Ballentine"}
    ElseIf ($User.UserName -eq "AREA52\1382931013N"){$Name = "Foster"}
    ElseIf ($User.UserName -eq "AREA52\1383257731N"){$Name = "Mowry"}
    ElseIf ($User.UserName -eq "AREA52\1470230947N"){$Name = "Lozada"}
    ElseIf ($User.UserName -eq "AREA52\1249051671N"){$Name = "Brown"}
    ElseIf ($User.UserName -eq "AREA52\1028801838N"){$Name = "Barnett"}
    ElseIf ($User.UserName -eq "AREA52\1366371229N"){$Name = "Cain"}
    ElseIf ($User.UserName -eq "AREA52\1252862141N"){$Name = "Simonds"}
    ElseIf ($User.UserName -eq "AREA52\1072361071"){$Name = "Ray"}
    ElseIf ($User.UserName -eq "AREA52\1082935297"){$Name = "Rick"}
    ElseIf ($User.UserName -eq "AREA52\1013110090N"){$Name = "Lewis"}
    ElseIf ($User.UserName -eq "AREA52\1116081047N"){$Name = "Carnall"}
    

$Number = Read-Host "Number"
$Phone = "$Number"
$Caller = Read-Host "Caller"
$Subject = Read-Host "Subject"

If (($User.UserName -eq "AREA52\1383807847N") -or 
    ($User.UserName -eq "AREA52\1253515879N") -or 
    ($User.UserName -eq "AREA52\1395576280N") -or 
    ($User.UserName -eq "AREA52\1382931013N") -or 
    ($User.UserName -eq "AREA52\1383257731N") -or
    ($User.UserName -eq "AREA52\1470230947N") -or
    ($User.UserName -eq "AREA52\1249051671N") -or
    ($User.UserName -eq "AREA52\1028801838N") -or
    ($User.UserName -eq "AREA52\1252862141N") -or
    ($User.UserName -eq "AREA52\1072361071") -or
    ($User.UserName -eq "AREA52\1082935297") -or
    ($User.UserName -eq "AREA52\1013110090N") -or
    ($User.UserName -eq "AREA52\1013110090N") -or
    ($User.UserName -eq "AREA52\1116081047N") -or
    ($User.UserName -eq "AREA52\1366371229N"))
    {$Message = "From: TSgt Simonds

You had a missed call from $Caller @ $Phone.

Subject: $Subject"
    Msg Console /Server:$Compname $Message
    Write-Host
    Write-Host "User Messaged: $Name"}
Else {Write-Host "The specified user is not logged on. Current user: $User" $User.UserName}
}

#===================================================================

Function ChangeBIOSPassword
{
$PWD = Read-Host "Current BIOS Password"
$NEWPWD = Read-Host "New BIOS Password (Cannot be blank)"
(Get-WmiObject -computername $Computer -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface).
SetBIOSSetting('Setup Password','<utf-16/>$NEWPWD','<utf-16/>$PWD')
}

#===================================================================

Function EnableBIOSComponent
{
$Device=Get-WmiObject -computername $Computer -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface
$Device.SetBIOSSetting("$Device","$EDselection")
}

#===================================================================

Function Enter-PSSession
{
Enter-PSSession -Computername $Computer
}

#===================================================================

Function ComputersPerBuilding
{
$BLDG = Read-Host "Building Number"
$domain = "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$objDomain = [adsi]("LDAP://" + $domain)
$search = New-Object System.DirectoryServices.DirectorySearcher
$search.SearchRoot = $objDomain
$search.Filter = "(&(objectClass=computer)(location=*BLDG: $BLDG*))"
$search.SearchScope = "Subtree"
$results = $search.FindAll()
ForEach($item in $results)
{
    $objComputer = $item.GetDirectoryEntry()
    $Name = $objComputer.cn
    Write-Host "$Name"
}
}

#===================================================================

Function ComputerAndUser
{
$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Computer"
$c.Cells.Item(1,2) = "User"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$Computers = Get-Content "C:\Users\1252862141.adm\Desktop\Scripts1\Enable_Local_Admin.txt"

ForEach ($Computer in $Computers)
{
    $c.Cells.Item($intRow,1) = $Computer
    
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Count 1 -Ea 0)
    {
        $User = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer
        If ($User.UserName -ne $null)
        {
            $EDI = $User.UserName.TrimStart("AREA52\")
            $UserInfo = Get-ADUser "$EDI" -Properties DisplayName
            $c.Cells.Item($intRow,2) = $UserInfo.DisplayName
        }
        Else
        {
            $c.Cells.Item($intRow,2) = "No user"
        }
    }
    Else
    {
        $c.Cells.Item($intRow,2) = "Offline"
        Write-Host "$Computer offline"
    }

    $intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()

$b.SaveAs("C:\work\Contacts_User_Info.xls")
}

#===================================================================


Function ResetAdmin
{


$in = Read-Host "Enter S for one PC or M for multiple"
    if(!(($in -eq "s") -or ($in -eq "m")))
      {
        $(Throw 'Input value must be (s) for single computer or (m) for multiple computers')
       }




If ($in -eq "s")
    {
        $computers = Read-Host "Enter the Computer Name" 
        # Update username / password as needed
        $username = "usaf_admin"
        $password = "zaq1XSW@zaq1XSW@"

            # Lists to store success / failed attempts
            $success = New-Object System.Collections.Generic.List[string]
            $failure = New-Object System.Collections.Generic.List[string]
        
                # Loop through each computer
                foreach ($computer in $computers) 
                    {# Attempt to change the password on the computer, ignoring any errors
                        try 
                        {([ADSI] "WinNT://$computer/$username").SetPassword("$password")} 
                            catch {}
                               # On success:
                                    if ($?) {$success.Add($computer)
                                    Write-Host "Success: $computer" -ForegroundColor Green}
                                
                                    # On failure:
                                        else {$failure.Add($computer)
                                        Write-Host "Failure: $computer" -ForegroundColor Red}
                                        
                               
                     }
     }
     
   

If ($in -eq "m")
    {



        $computers = get-content "c:\users\1383807847.adm\desktop\scripts\computer.txt"
        # Update username / password as needed
        $username = "usaf_admin"
        $password = "zaq1XSW@zaq1XSW@"

            # Lists to store success / failed attempts
            $success = New-Object System.Collections.Generic.List[string]
            $failure = New-Object System.Collections.Generic.List[string]

                # Loop through each computer
                foreach ($computer in $computers) 
                    {# Attempt to change the password on the computer, ignoring any errors
                        try 
                        {([ADSI] "WinNT://$computer/$username").SetPassword("$password")} 
                        catch {}
                            # On success:
                                if ($?) {$success.Add($computer)
                                Write-Host "Success: $computer" -ForegroundColor Green}
    
                            # On failure:
                                else {$failure.Add($computer)
                            Write-Host "Failure: $computer" -ForegroundColor Red}
                    }
}


}

#===================================================================



Do
{
    Write-Host " "
    Write-Host -ForegroundColor Green "0 - Cls"
    Write-Host -ForegroundColor Green "1 - Send Message"
    Write-Host -ForegroundColor Green "2 - Lock Computer"
    Write-Host -ForegroundColor Green "3 - Logoff User"
    Write-Host -ForegroundColor Green "4 - Reboot Computer"
    Write-Host -ForegroundColor Green "5 - Enable Local Admin"
    Write-Host -ForegroundColor Green "6 - Missed Call"
    Write-Host -ForegroundColor Green "7 - Change BIOS Password (Don't use)"
    Write-Host -ForegroundColor Green "8 - Enable/Disable BIOS Component"
    Write-Host -ForegroundColor Green "9 - Enter-PSSession"
    Write-Host -ForegroundColor Green "10 - Computer & Building"
    Write-Host -ForegroundColor Green "11 - Computer & User"
    Write-Host -ForegroundColor Green "12 - Reset Local Admin"
    Write-Host -ForegroundColor Green "13 - Exit"
    Write-Host " "

    $Ans = Read-Host "Pick something why don't ya?"
    
    If ($Ans -eq 0)
    {
        cls
    }
    If ($Ans -eq 1)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        SendMessage
    }
    If ($Ans -eq 2)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        LockComputer
    }
    If ($Ans -eq 3)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        LogoffUser
    }
    If ($Ans -eq 4)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        RebootComputer
    }
    If ($Ans -eq 5)
    {
        EnableLocalAdmin
    }
    If ($Ans -eq 6)
    {
        MissedCall
    } 
    If ($Ans -eq 7)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        ChangeBIOSPassword
    }
    If ($Ans -eq 8)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        $EorD = Read-Host "Would you like to Enable (e) or Disable (d) a BIOS component?"
            
            If ($EorD -eq "e")
                {
                    $EDselection = "Enable"
                }
            
            If ($EorD -eq "d")
                {
                    $EDselection = "Disable"
                }
        
        Write-Host "1 - CD-ROM Boot"
        Write-Host "2 - Network (PXE) Boot"
        Write-Host "3 - Power On When Lid is Opened"
        Write-Host "4 - NumLock on at boot"
        Write-Host "5 - Legacy Support Disable and Secure Boot Enable"
        Write-Host "6 - Legacy Support Enable and Secure Boot Disable"
        Write-Host "7 - Legacy Support Disable and Secure Boot Disable"
        Write-Host "8 - Audio Device"
        Write-Host "9 - Integrated Microphone"
        Write-Host "10 - Internal Speakers"
        Write-Host "11 - Headphone Output"
        Write-Host "12 - Lock Wireless Button"
        Write-Host "13 - Wireless Network Device (WLAN)"
        Write-Host "14 - Bluetooth"
        Write-Host "15 - Lan / WLAN Auto Switching"
        Write-Host "16 - Wake on WLAN"
        Write-Host "17 - Wake on LAN in Battery Mode"
        Write-Host "18 - Fan Always on while on AC Power"
        Write-Host "19 - Boost Converter"
        Write-Host "20 - Integrated Camera"
        Write-Host "21 - Fingerprint Device"
        Write-Host "22 - Fingerprint Reset on Reboot"
        Write-Host "23 - Prompt for Admin password on F9 (Boot Menu)"
        Write-Host "24 - Prompt for Admin password on F11 (System Recovery"
        Write-Host "25 - Prompt for Admin password on F12 (Network Boot)" 
        $Ans2 = Read-Host "Pick one"   
            If ($Ans2 -eq 1)
                 {
                    $Device = "CD-ROM Boot"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 2)
                 {
                    $Device = "Network (PXE) Boot"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 3)
                 {
                    $Device = "Power On When Lid is Opened"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 4)
                 {
                    $Device = "NumLock on at boot"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 5)
                 {
                    $Device = "Legacy Support Disable and Secure Boot Enable"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 6)
                 {
                    $Device = "Legacy Support Enable and Secure Boot Disable"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 7)
                 {
                    $Device = "Legacy Support Disable and Secure Boot Disable"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 8)
                 {
                    $Device = "Audio Device"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 9)
                 {
                    $Device = "Integrated Microphone"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 10)
                 {
                    $Device = "Internal Speakers"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 11)
                 {
                    $Device = "Headphone Output"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 12)
                 {
                    $Device = "Lock Wireless Button"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 13)
                 {
                    $Device = "Wireless Network Device (WLAN)"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 14)
                 {
                    $Device = "Bluetooth"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 15)
                 {
                    $Device = "Lan / WLAN Auto Switching"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 16)
                 {
                    $Device = "Wake on WLAN"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 17)
                 {
                    $Device = "Wake on LAN in Battery Mode"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 18)
                 {
                    $Device = "Fan Always on while on AC Power"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 19)
                 {
                    $Device = "Boost Converter"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 20)
                 {
                    $Device = "Integrated Camera"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 21)
                 {
                    $Device = "Fingerprint Device"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 22)
                 {
                    $Device = "Fingerprint Reset on Reboot"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 23)
                 {
                    $Device = "Prompt for Admin password on F9 (Boot Menu)"
                    EnableBIOSComponent
                 }
            If ($Ans2 -eq 24)
                 {
                    $Device = "Prompt for Admin password on F11 (System Recovery)"
                    EnableBIOSComponent
                 } 
            If ($Ans2 -eq 25)
                 {
                    $Device = "Prompt for Admin password on F12 (Network Boot)"
                    EnableBIOSComponent
                 }                                                      
    }
    If ($Ans -eq 9)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        Enter-PSSession
    }
    If ($Ans -eq 10)
    {
        ComputersPerBuilding
    }
    If ($Ans -eq 11)
    {
        ComputerAndUser
    }
    If ($Ans -eq 12)
    {
        ResetAdmin
    }
    
}
Until ($Ans -eq 13)