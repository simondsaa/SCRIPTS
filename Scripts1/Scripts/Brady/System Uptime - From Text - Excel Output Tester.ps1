$NetAdmin = "325CS.SCOO.NetAdmin@us.af.mil"
$DayLimit = 0
$Computers = Get-Content "C:\Users\timothy.brady\Desktop\Servers.txt"
$Body = "$User,

$strComputer has not been rebooted in $Days days, please reboot this server as soon as possible to ensure patches are being installed.

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"

ForEach ($strComputer in $Computers)
{
    If (Test-Connection $strComputer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        $LastBootTime = (Get-WmiObject Win32_OperatingSystem -cn $strComputer -ErrorAction SilentlyContinue).LastBootUpTime
        $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
        $Days = $Uptime.days
            
        If ($Days -gt $DayLimit)
        {
            If ($strComputer -eq "TYNCESAAPPD2P2" -or "TYNCESADBPD2P2")
            {
                $User = "Mr. Drake"
                #$UserMail = "william.drake.3.ctr@us.af.mil"
            }

            ElseIf ($strComputer -eq "52XLWU-IS-001p" -or "52XLWU-IS-001p" -or "52XLWU-IS-001p" -or "52XLWU-IS-001p" -or "52XLWU-IS-001p")
            {
                $User = "Mr. Stringer"
                $UserMail = "lonnie.stringer.3@us.af.mil"
                Send-MailMessage -From $NetAdmin -To "timothy.brady.11@us.af.mil" -Priority High -Subject "Test" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
            }
            
            ElseIf ($strComputer -eq "XLWU-SM-001" -or "XLWU-SM-002")
            {
                $User = "SrA Brady"
                $UserMail = "timothy.brady.11@us.af.mil"
                Send-MailMessage -From $NetAdmin -To "timothy.brady.11@us.af.mil" -Priority High -Subject "Test" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
            }
        
            ElseIf ($strComputer -eq "TYNAFRLAP60402" -or "TYNAFRLAP60502" -or "TYNAFRLAP60602" -or "TYNAFRLAPV0102")
            {
                $User = "Mr. Sindt"
                #$UserMail = "caleb.sindt.1.ctr@us.af.mil"
            }
    
            ElseIf ($strComputer -eq "TYNCESAAPMSC02" -or "52XLWU-FS-011" -or "TYNCESAAPSPD02" -or "TYNCESAAPIBM02" -or "TYNCESAAPSRP02" -or "TYNCESAFSEDIE03" -or "TYNCESAAPSQL02" -or "TYNCESAAPWAP02")
            {
                $User = "Mr. Malott"
                #$UserMail = "christopher.malott.2.ctr@us.af.mil"
            }
    
            ElseIf ($strComputer -eq "TYNSUPPLY02")
            {
                $User = "Mr. Vachon"
                #$UserMail = "carl.vachon.ctr@us.af.mil"
            }
            
            ElseIf ($strComputer -eq "TYNDS2XFS01")
            {
                $User = "Mr. Grape"
                #$UserMail = "eric.grape.ctr@us.af.mil"
            }
        
            ElseIf ($strComputer -eq "TYNMDGSGSV00603" -or "TYNMDGSGSV00403" -or "TYNMDGSGSVSN103" -or "TYNMDGSGSVSN203" -or "52XLWU-AS-002p" -or "52XLWU-MD-001" -or "TYNMDGSGSV00503")
            {
                $User = "Mr. Charles"
                #$UserMail = "jason.charles.ctr@us.af.mil"
            }        

            ElseIf ($strComputer -eq "52XLWU-FS-010" -or "TYNFS10")
            {
                $User = "Mr. Gardner"
                #$UserMail = "jerry.gardner@us.af.mil"
            }
    
            ElseIf ($strComputer -eq "TYNAP001P1" -or "TYNAP002P1" -or "52XLWU-FS-GDBp")
            {
                $User = "Mr. Barnes"
                #$UserMail = "kenneth.barnes@us.af.mil"
            }
    
            ElseIf ($strComputer -eq "52GWDD-AS01-01" -or "52GWDD-BS01-01" -or "52GWDD-CL01-01" -or "52GWDD-CL02-02" -or "52GWDD-SQL01-01" -or "52GWDD-SQL02-02")
            {
                $User = "Mr. Smith"
                #$UserMail = "robert.smith.209.ctr@us.af.mil"
            }
    
            ElseIf ($strComputer -eq "52XLWU-DB05-01" -or "52XLWU-DB06-01" -or "52XLWU-CL05-01" -or "52XLWU-CL06-01" -or "52XLWU-DB07-01")
            {
                $User = "Mr. Pandullo"
                #$UserMail = "ronald.pandullo.1@us.af.mil"
            }
    
            ElseIf ($strComputer -eq "TYNCONSASFSVR" -or "TYNCONSDBSERVER")
            {
                $User = "Mr. Bloedow"
                #$UserMail = "david.kigerl.2@us.af.mil"
            }
    
            ElseIf ($strComputer -eq "TYNLRDLGFMDP3")
            {
                $User = "Mr. Nelson"
                #$UserMail = "timothy.nelson.18.ct@us.af.mil"
            }
    
            ElseIf ($strComputer -eq "TYNOGAPPEX0103" -or "TYNOOGAPPEX203")
            {
                $User = "Mr. Newman"
                #$UserMail = "william.newman.16@us.af.mil"
            }

            ElseIf ($strComputer -eq "TYNOOSSAAP0103")
            {
                $User = "Ms. Taylor"
                #$UserMail = "deborah.taylor.7@us.af.mil"
            }

            ElseIf ($strComputer -eq "JETKPAM")
            {
                $User = "MSgt Boling"
                #$UserMail = "bruce.boling@us.af.mil"
            }
        }
    }
}
            
$c = New-Object -Comobject wscript.shell
$b = $c.popup("The script has completed",0,"Complete",80)