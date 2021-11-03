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
        $Time = 3600
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

$user = Read-Host "User Name"

if($a -eq "e")

          {
                $password = Read-Host "New Password"
        
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
Do
{
    Cls
    Write-Host " "
    Write-Host "1 - Send Message"
    Write-Host "2 - Lock Computer"
    Write-Host "3 - Logoff User"
    Write-Host "4 - Reboot Computer"
    Write-Host "5 - Enable Local Admin"
    Write-Host "6 - Exit"
    Write-Host " "

    $Ans = Read-Host "Make Selection"
    
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
        Write-Host
        $Computer = Read-Host "Computer"
        RebootComputer
    }
}
Until ($Ans -eq 6)