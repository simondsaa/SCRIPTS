$NetAdmin = "325CS.SCOO.NetAdmin@us.af.mil"
$DayLimit = 7
$Path = "C:\Users\timothy.brady\Music"
$Message = "The following server(s) have not been rebooted within $DayLimit days, please reboot them at your earliest convenience."
$Notice = "*This is an automatically generated email from PowerShell, if you have questions please reply directly to this email*"

Function SCOO
{
    $User = "SCOO"
    $UserMail = "325CS.SCOO.NetAdmin@us.af.mil"
    $Computers = Get-Content "$Path\SCOO.txt"

    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            $LastBootTime = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).LastBootUpTime
            $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
            $Days = $Uptime.days
        }
        If ($Days -gt $DayLimit)
        {
            $Table += "$Computer   -   $Days days
"
        }
    }
    If ($Days -gt $DayLimit)
    {
        $Body = "$User,

$Message

$Table
$Notice

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"
        Send-MailMessage -From $NetAdmin -To $UserMail -Bcc $NetAdmin -Priority High -Subject "Server Uptime - $User" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
    }
}
SCOO

Function Valdez
{
    $User = "SSgt Valdez"
    $UserMail = "andrew.valdez@us.af.mil"
    $Computers = Get-Content "$Path\Valdez.txt"

    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            $LastBootTime = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).LastBootUpTime
            $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
            $Days = $Uptime.days
        }
        If ($Days -gt $DayLimit)
        {
            $Table += "$Computer   -   $Days days
"
        }
    }
    If ($Days -gt $DayLimit)
    {
        $Body = "$User,

$Message

$Table
$Notice

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"
        Send-MailMessage -From $NetAdmin -To $UserMail -Bcc $NetAdmin -Priority High -Subject "Server Uptime - $User" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
    }
}
Valdez

Function Sindt
{
    $User = "Mr. Sindt"
    $UserMail = "caleb.sindt.1.ctr@us.af.mil"
    $Computers = Get-Content "$Path\Sindt.txt"

    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            $LastBootTime = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).LastBootUpTime
            $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
            $Days = $Uptime.days
        }
        If ($Days -gt $DayLimit)
        {
            $Table += "$Computer   -   $Days days
"
        }
    }
    If ($Days -gt $DayLimit)
    {
        $Body = "$User,

$Message

$Table
$Notice

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"
        Send-MailMessage -From $NetAdmin -To $UserMail -Bcc $NetAdmin -Priority High -Subject "Server Uptime - $User" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
    }
}
Sindt

Function Malott
{
    $User = "Mr. Malott"
    $UserMail = "christopher.malott.2.ctr@us.af.mil"
    $Computers = Get-Content "$Path\Malott.txt"

    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            $LastBootTime = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).LastBootUpTime
            $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
            $Days = $Uptime.days
        }
        If ($Days -gt $DayLimit)
        {
            $Table += "$Computer   -   $Days days
"
        }
    }
    If ($Days -gt $DayLimit)
    {
        $Body = "$User,

$Message

$Table
$Notice

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"
        Send-MailMessage -From $NetAdmin -To $UserMail -Bcc $NetAdmin -Priority High -Subject "Server Uptime - $User" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
    }
}
Malott

Function Charles
{
    $User = "Mr. Charles"
    $UserMail = "jason.charles.ctr@us.af.mil"
    $Computers = Get-Content "$Path\Charles.txt"

    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            $LastBootTime = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).LastBootUpTime
            $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
            $Days = $Uptime.days
        }
        If ($Days -gt $DayLimit)
        {
            $Table += "$Computer   -   $Days days
"
        }
    }
    If ($Days -gt $DayLimit)
    {
        $Body = "$User,

$Message

$Table
$Notice

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"
        Send-MailMessage -From $NetAdmin -To $UserMail -Bcc $NetAdmin -Priority High -Subject "Server Uptime - $User" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
    }
}
Charles

Function Barnes
{
    $User = "Mr. Barnes"
    $UserMail = "kenneth.barnes@us.af.mil"
    $Computers = Get-Content "$Path\Barnes.txt"

    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            $LastBootTime = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).LastBootUpTime
            $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
            $Days = $Uptime.days
        }
        If ($Days -gt $DayLimit)
        {
            $Table += "$Computer   -   $Days days
"
        }
    }
    If ($Days -gt $DayLimit)
    {
        $Body = "$User,

$Message

$Table
$Notice

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"
        Send-MailMessage -From $NetAdmin -To $UserMail -Cc "jim.whitcomb.1@us.af.mil" -Bcc $NetAdmin -Priority High -Subject "Server Uptime - $User" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
    }
}
Barnes

Function Bloedow
{
    $User = "Mr. Bloedow"
    $UserMail = "ryan.bloedow@us.af.mil"
    $Computers = Get-Content "$Path\Bloedow.txt"

    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            $LastBootTime = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).LastBootUpTime
            $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
            $Days = $Uptime.days
        }
        If ($Days -gt $DayLimit)
        {
            $Table += "$Computer   -   $Days days
"
        }
    }
    If ($Days -gt $DayLimit)
    {
        $Body = "$User,

$Message

$Table
$Notice

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"
        Send-MailMessage -From $NetAdmin -To $UserMail -Bcc $NetAdmin -Priority High -Subject "Server Uptime - $User" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
    }
}
Bloedow

Function Taylor
{
    $User = "Mrs. Taylor"
    $UserMail = "deborah.taylor.7@us.af.mil"
    $Computers = Get-Content "$Path\Taylor.txt"

    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            $LastBootTime = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).LastBootUpTime
            $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
            $Days = $Uptime.days
        }
        If ($Days -gt $DayLimit)
        {
            $Table += "$Computer   -   $Days days
"
        }
    }
    If ($Days -gt $DayLimit)
    {
        $Body = "$User,

$Message

$Table
$Notice

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"
        Send-MailMessage -From $NetAdmin -To $UserMail -Bcc $NetAdmin -Priority High -Subject "Server Uptime - $User" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
    }
}
Taylor

Function Boling
{
    $User = "SMSgt Boling"
    $UserMail = "bruce.boling@us.af.mil"
    $Computers = Get-Content "$Path\Boling.txt"

    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            $LastBootTime = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).LastBootUpTime
            $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
            $Days = $Uptime.days
        }
        If ($Days -gt $DayLimit)
        {
            $Table += "$Computer   -   $Days days
"
        }
    }
    If ($Days -gt $DayLimit)
    {
        $Body = "$User,

$Message

$Table
$Notice

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"
        Send-MailMessage -From $NetAdmin -To $UserMail -Bcc $NetAdmin -Priority High -Subject "Server Uptime - $User" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
    }
}
Boling

Function Pandullo
{
    $User = "Mr. Pandullo"
    $UserMail = "ronald.pandullo.1@us.af.mil"
    $Computers = Get-Content "$Path\Pandullo.txt"

    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            $LastBootTime = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).LastBootUpTime
            $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
            $Days = $Uptime.days
        }
        If ($Days -gt $DayLimit)
        {
            $Table += "$Computer   -   $Days days
"
        }
    }
    If ($Days -gt $DayLimit)
    {
        $Body = "$User,

$Message

$Table
$Notice

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"
        Send-MailMessage -From $NetAdmin -To $UserMail -Bcc $NetAdmin -Priority High -Subject "Server Uptime - $User" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
    }
}
Pandullo

Function Thompson
{
    $User = "Mr. Thompson"
    $UserMail = "james.thompson.90@us.af.mil"
    $Computers = Get-Content "$Path\Thompson.txt"

    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            $LastBootTime = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).LastBootUpTime
            $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
            $Days = $Uptime.days
        }
        If ($Days -gt $DayLimit)
        {
            $Table += "$Computer   -   $Days days
"
        }
    }
    If ($Days -gt $DayLimit)
    {
        $Body = "$User,

$Message

$Table
$Notice

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"
        Send-MailMessage -From $NetAdmin -To $UserMail -Bcc $NetAdmin -Priority High -Subject "Server Uptime - $User" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
    }
}
Thompson

Function Stieber
{
    $User = "Ms. Stieber"
    $UserMail = "heather.stieber.ctr@us.af.mil"
    $Computers = Get-Content "$Path\Stieber.txt"

    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            $LastBootTime = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).LastBootUpTime
            $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
            $Days = $Uptime.days
        }
        If ($Days -gt $DayLimit)
        {
            $Table += "$Computer   -   $Days days
"
        }
    }
    If ($Days -gt $DayLimit)
    {
        $Body = "$User,

$Message

$Table
$Notice

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"
        Send-MailMessage -From $NetAdmin -To $UserMail -Cc "kenneth.barnes@us.af.mil" -Bcc $NetAdmin -Priority High -Subject "Server Uptime - $User" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
    }
}
#Stieber