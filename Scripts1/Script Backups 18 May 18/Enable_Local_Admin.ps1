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

       $(Write-Error 'This PC has either died or has one foot out the door. Good luck.')

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