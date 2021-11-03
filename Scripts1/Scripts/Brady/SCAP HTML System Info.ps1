$LocalUser = Get-WmiObject Win32_ComputerSystem
$LocalEDI = $LocalUser.UserName.TrimStart("AREA52\")
$LocalName = (Get-ADUser "$LocalEDI" -Properties DisplayName).DisplayName

$Computer = Read-Host "Computer Name"

$Date = Get-Date -UFormat "%d-%b-%g %H:%M"

$ProgHTML = $null

If (Test-Path "C:\Users\$LocalEDI\Desktop\Remote.bat")
{
    Remove-Item -Path C:\Users\$LocalEDI\Desktop\Remote.bat -Force
}

If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
{
    $NetInfo = Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled = $true" -ComputerName $Computer -ErrorAction SilentlyContinue | Where-Object {$_.IPAddress -like "131.55*"}
    $NIC = $NetInfo.Description
    $IP = $NetInfo.IPAddress
    $MAC = $NetInfo.MACAddress

    $SysInfo = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer -ErrorAction SilentlyContinue
    $SysName = $SysInfo.Name
    $RAM = [Math]::Round(($SysInfo.TotalPhysicalMemory) / 1048576, 0)
    $Manufacturer = $SysInfo.Manufacturer
    $Domain = $SysInfo.Domain
    $Model = $SysInfo.Model
    $Bit = $SysInfo.SystemType

    $OSInfo = Get-Wmiobject Win32_OperatingSystem -ComputerName $Computer -ErrorAction SilentlyContinue
    $OS = $OSInfo.Caption
    $SP = $OSInfo.ServicePackMajorVersion
    $SDC = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation').GetValue('Model')
    $sysuptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($OSInfo.LastBootUpTime)
    $UpDays = $sysuptime.days
    $UpHours = $sysuptime.hours
    $UpMins = $sysuptime.minutes
        
    $Serial = (Get-Wmiobject Win32_Bios -ComputerName $Computer -ErrorAction SilentlyContinue).SerialNumber
    $CPU = (Get-WmiObject Win32_Processor -ComputerName $Computer -ErrorAction SilentlyContinue).Name

    $Profiles = Get-ChildItem \\$Computer\C$\Users
    $AdminProf = 0
    ForEach ($Profile in $Profiles)
    {
        If ($Profile -like "*.adm")
        {
            $AdminProf += 1
        }
    }
    $ProfCount = $Profiles.Count

    If ($OSInfo.OSArchitecture -eq "64-bit"){$RegPath = "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"}
    ElseIf ($OSInfo.OSArchitecture -eq "32-bit"){$RegPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall"}        
    $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer)
    $RegKey = $Reg.OpenSubKey($RegPath)
    $SubKeys = $RegKey.GetSubKeyNames()
    $Array = @()
    ForEach($Key in $SubKeys)
    {
        $ThisKey = $RegPath+"\"+$Key 
        $ThisSubKey = $Reg.OpenSubKey($ThisKey)
        $obj = New-Object PSObject
        $obj | Add-Member -Force -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
        $obj | Add-Member -Force -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
        $obj | Add-Member -Force -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
        $obj | Add-Member -Force -MemberType NoteProperty -Name "InstallDate" -Value $($thisSubKey.GetValue("InstallDate"))
        $obj | Add-Member -Force -MemberType NoteProperty -Name "HelpLink" -Value $($thisSubKey.GetValue("HelpLink"))
        $Array += $obj
    }
    $Progs = $Array | Where-Object {($_.Publisher -ne $null) -and ($_.DisplayName -ne $null)} | Sort-Object "DisplayName" #| ConvertTo-HTML -Body $Body

    ForEach ($Prog in $Progs)
    {
        $ProgName = $Prog.DisplayName
        $ProgVer = $Prog.DisplayVersion
        $ProgPub = $Prog.Publisher
        $ProgInst = $Prog.InstallDate
        $ProgHelp = $Prog.HelpLink
        $ProgHTML += "<tr>
        <td class='first'>
        <DT>$ProgName</DT>
        </td>
        <td class='first'>
        <DT>$ProgVer</DT>
        </td>
        <td class='first'>
        <DT>$ProgInst</DT>
        </td>
        <td class='first'>
        <DT>$ProgPub</DT>
        </td>
        <td class='first'>
        <A HREF=$ProgHelp target='_blank'>$ProgHelp</A>
        </td>
        </tr>"
    }
       
    Do
    {
        $RegCheck = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SYSTEM\CurrentControlSet\Control\Terminal Server').GetValue('AllowRemoteRPC')
        If ($RegCheck -ne "1")
        {
            Write-Host "Running Registry fix..."
            Start-Sleep -Seconds 1
            REG ADD "\\$Computer\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v AllowRemoteRPC /t REG_DWORD /d 1 /f
            Start-Sleep -Seconds 1
        }        
        Try 
        {
            $User = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer
            If ($User.UserName -ne $null)
            {
                $EDI = $User.UserName.TrimStart("AREA52\")
                $UserInfo = Get-ADUser "$EDI" -Properties DisplayName, City, gigID, EmailAddress, extensionAttribute5, mDBOverHardQuotaLimit, LockedOut, Enabled, OfficePhone, MemberOf -ErrorAction SilentlyContinue
                $MailSize = ($UserInfo.mDBOverHardQuotaLimit/1024)
                $Name = $UserInfo.DisplayName
                $Pre = $UserInfo.SamAccountName
                $Base = $UserInfo.City
                $Email = $UserInfo.EmailAddress
                $Cat = $UserInfo.extensionAttribute5
                $Locked = $UserInfo.LockedOut
                $Enabled = $UserInfo.Enabled
                $Number = $UserInfo.OfficePhone
            }
            Else
            {
                Write-Host "No user logged on." -ForegroundColor Yellow
                $Name = "No user logged on."
                $MailSize = $null
                $Pre = $null
                $Base = $null
                $Email = $null
                $Cat = $null
                $Locked = $null
                $Enabled = $null
                $Number = $null
            }
        }
        Catch
        {
            Write-Host "Failed to get User information"
        }
    }
    While ($RegCheck -ne "1")

    New-Item -Path C:\Users\$LocalEDI\Desktop -Name Remote.bat -ItemType file -Force | Out-Null
    Add-Content -Path C:\Users\$LocalEDI\Desktop\Remote.bat -Value "mstsc /v:$SysName /f" -Force
    Add-Content -Path C:\Users\$LocalEDI\Desktop\Remote.bat -Value "Exit" -Force
}
Else
{
    Write-Host $Computer "offline"
    Exit
}

$Head = "<head>
<title>System Report</title>
<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />
<style type='text/css'>
body
{
background-color: #EDEDED;
margin: 25px 25px 25px 25px;
padding: 0px 0px 0px 0px;
font-size: 12px;
text-align: left;
font-family: arial, helvetica, verdana, sans-serif;
color: #3A5064;
font-weight: normal;
}

hr
{
border: 0;
height: 1px;
border-bottom: 1px dotted #B3B3B3;
}

table
{
background-color: transparent;
margin: 0px 0px 0px 0px;
padding: 0px 0px 0px 0px;
border-top: 1px solid #B3B3B3; 
border-left: 1px solid #B3B3B3; 
font-size: 12px;
}

td
{
padding: 5px 5px 5px 5px;
border-bottom: 1px solid #B3B3B3;
border-right: 1px solid #B3B3B3; 
text-align: left;
vertical-align: top;
}

td.first
{
padding: 5px 5px 5px 5px;
border-bottom: 1px solid #B3B3B3;
border-right: 1px solid #B3B3B3; 
text-align: left;
vertical-align: top;
white-space: nowrap;
}

table.siteSummary
{
background-color: transparent;
margin: 0px 0px 0px 0px;
padding: 0px 0px 0px 0px;
border-top: 1px solid #B3B3B3; 
border-left: 1px solid #B3B3B3; 
font-size: 12px;
}

table.siteSummary th
{
background-color: #0F6FBB;
padding: 5px 5px 5px 5px;
border-bottom: 1px solid #B3B3B3;
border-right: 1px solid #B3B3B3; 
text-align: left;
vertical-align: middle;
color: #FFFFFF;
font-weight: bold;
}

table.siteSummary th.centered
{
background-color: #0F6FBB;
padding: 5px 5px 5px 5px;
border-bottom: 1px solid #B3B3B3;
border-right: 1px solid #B3B3B3; 
text-align: center;
vertical-align: middle;
color: #FFFFFF;
font-weight: bold;
}

table.siteSummary td
{
padding: 5px 5px 5px 5px;
border-bottom: 1px solid #B3B3B3;
border-right: 1px solid #B3B3B3; 
text-align: left;
vertical-align: top;
white-space: nowrap;
}

table.siteSummary td.centered
{
padding: 5px 5px 5px 5px;
border-bottom: 1px solid #B3B3B3;
border-right: 1px solid #B3B3B3; 
text-align: center;
vertical-align: top;
white-space: nowrap;
}

table.computerList
{
background-color: transparent;
margin: 0px 0px 0px 0px;
padding: 0px 0px 0px 0px;
border-top: 1px solid #B3B3B3; 
border-left: 1px solid #B3B3B3; 
font-size: 12px;
}

table.computerList th
{
background-color: #0F6FBB;
padding: 5px 5px 5px 5px;
border-bottom: 1px solid #B3B3B3;
border-right: 1px solid #B3B3B3; 
text-align: center;
vertical-align: middle;
color: #FFFFFF;
font-weight: bold;
}

table.computerList td
{
padding: 5px 5px 5px 5px;
border-bottom: 1px solid #B3B3B3;
border-right: 1px solid #B3B3B3; 
text-align: center;
vertical-align: top;
white-space: nowrap;
}

table.score
{
background-color: transparent;
margin: 0px 0px 0px 0px;
padding: 0px 0px 0px 0px;
border: none;
font-size: 12px;
text-align: left;
}

table.score td.first
{
padding: 0px 15px 0px 15px;
border: none;
text-align: center;
vertical-align: middle;
white-space: nowrap;
font-size: 86px;
}

table.score td.left
{
padding: 0px 15px 0px 0px;
border: none;
text-align: center;
vertical-align: middle;
white-space: nowrap;
}

table.score td
{
padding: 0px 15px 0px 15px;
border: none;
border-left: 1px dotted #B3B3B3;
text-align: center;
vertical-align: middle;
}

table.score td.label
{
padding: 0px 5px 0px 10px;
border: none;
text-align: right;
vertical-align: top;
white-space: nowrap;
}

table.score td.value
{
padding: 0px 10px 0px 5px;
border: none;
text-align: left;
vertical-align: top;
white-space: nowrap;
}

table.score td.labelBold
{
padding: 0px 5px 0px 10px;
border: none;
text-align: right;
vertical-align: top;
white-space: nowrap;
font-weight: bold;
}

table.score td.valueBold
{
padding: 0px 10px 0px 5px;
border: none;
text-align: left;
vertical-align: top;
white-space: nowrap;
font-weight: bold;
}

table.transparent
{
background-color: transparent;
margin: 0px 0px 0px 0px;
padding: 0px 0px 0px 0px;
border: none;
font-size: 12px;
}

table.transparent td
{
padding: 0px 5px 0px 5px;
border: none;
text-align: left;
vertical-align: top;
}

table.transparent td.first
{
padding: 0px 5px 0px 5px;
border: none;
text-align: right;
vertical-align: top;
white-space: nowrap;
}

h1
{
padding: 25px 0px 0px 0px;
font-size: 30px;
font-weight: bold;
border-bottom: 2px solid #B3B3B3;
}

h2
{
padding: 25px 0px 0px 0px;
font-size: 16px;
font-weight: bold;
}

b
{
font-size: 12px;
font-weight: bold;
}		

ul
{
background-color: transparent;
margin: 0px 20px 0px 20px;
padding: 0px 0px 0px 0px;
color: #3A5064;
}

.alt
{
background-color: #E7F3FD;
}

.pass
{
color: #0080FF;
}

.fail
{
color: #FF0000;
}

.nostatus
{
color: #3A5064;
}

.nowrap
{
white-space: nowrap;
}

li a:link {text-decoration: none; color: #0080FF;} 
li a:visited {text-decoration: none; color: #0080FF;}
li a:hover {text-decoration: underline; color: #0080FF;} 
li a:active {text-decoration: underline; color: #0080FF;}	

li a.fail:link {text-decoration: none; color: #FF0000;} 
li a.fail:visited {text-decoration: none; color: #FF0000;}
li a.fail:hover {text-decoration: underline; color: #FF0000;} 
li a.fail:active {text-decoration: underline; color: #FF0000;}	

li a.nostatus:link {text-decoration: none; color: #3A5064;} 
li a.nostatus:visited {text-decoration: none; color: #3A5064;}
li a.nostatus:hover {text-decoration: underline; color: #3A5064;} 
li a.nostatus:active {text-decoration: underline; color: #3A5064;}	

#header
{
background-color: #0F6FBB;
margin: 0px 0px 0px 0px;
padding: 25px 25px 25px 25px;
border-top: 1px solid #B3B3B3;
border-left: 1px solid #B3B3B3;
border-right: 1px solid #B3B3B3;
color: #95CFFB;
font-size: 12px;
font-weight: bold;
}

#header h1
{
color: #FFFFFF;
font-size: 30px;
font-weight: bold;
padding: 0px 0px 0px 0px;
margin: 0px 0px 0px 0px;
border: none;
}

#navigation
{
background-color: #F7F7F7;
margin: 0px 0px 0px 0px;
padding: 5px 25px 5px 25px;
border-top: 1px solid #B3B3B3;
border-left: 1px solid #B3B3B3;
border-right: 1px solid #B3B3B3;
color: #737373;
font-weight: bold;
}

#navigation a:link {text-decoration: none; color: #3A5064;} 
#navigation a:visited {text-decoration: none; color: #3A5064;}
#navigation a:hover {text-decoration: underline; color: #3A5064;} 
#navigation a:active {text-decoration: underline; color: #3A5064;}	

#container
{
background-color: transparent;
margin: 0px 0px 0px 0px;
padding: 0px 0px 0px 0px;
border: 1px solid #B3B3B3;
}		

#content
{
background-color: #FFFFFF;
margin: 0px 0px 0px 0px;
padding: 10px 25px 25px 25px;
}

#footer
{
background-color: transparent;
margin: 0px 0px 0px 0px;
padding: 5px 25px 5px 25px;
color: #808080;
font-size: 12px;	
text-align: center;
}
</style>
</head>"
$Body = "<body>
<div id='header'>
<h1>Full System Report: $SysName</h1>Report Generated: $Date by $LocalName</div>
<div id='navigation'><a href='#systemInformation'>System Information</a> | <a href='#userInformation'>User Information</a> | <a href='#installedPrograms'>Installed Programs</a></div>
<div id='container'>
<div id='content'>
<a name='systemInformation'><h1>System Information</h1></a>
<table border='0' cellspacing='0' align='center'>
<tr class='alt'>
<td class='first'>Target:</td>
<td width='100%'><A HREF='C:\Users\$LocalEDI\Desktop\Remote.bat'>$SysName</A></td>
</tr>
<tr>
<td class='first'>Operating System:</td>
<td>$OS SP $SP</td>
</tr>
<tr class='alt'>
<td class='first'>SDC Version:</td>
<td>$SDC</td>
</tr>
<tr>
<td class='first'>System Bit:</td>
<td>$Bit</td>
</tr>
<tr class='alt'>
<td class='first'>Domain:</td>
<td>$Domain</td>
</tr>
<tr>
<td class='first'>Processor:</td>
<td>$CPU</td>
</tr>
<tr class='alt'>
<td class='first'>Physical Memory:</td>
<td>$RAM MB</td>
</tr>
<tr>
<td class='first'>Manufacturer:</td>
<td>$Manufacturer</td>
</tr>
<tr class='alt'>
<td class='first'>Model:</td>
<td>$Model</td>
</tr>
<tr>
<td class='first'>Serial Number:</td>
<td>$Serial</td>
</tr>
<tr class='alt'>
<td class='first'>Interfaces:</td>
<td>
<ul>
<li>$NIC</li>
<ul>
<li>$IP</li>
<li>$MAC</li>
</ul>
</ul>
</td>
</tr>
<tr>
<td class='first'>Uptime:</td>
<td>$UpDays day(s) $UpHours hours $UpMins mins</td>
</tr>
<tr class='alt'>
<td class='first'>Profiles:</td>
<td>$ProfCount total   |   $AdminProf admin profile(s)</td>
</tr>
</table>
<a name='userInformation'><h1>User Information</h1></a>
<table border='0' cellspacing='0' align='center'>
<tr class='alt'>
<td class='first'>User Name:</td>
<td width='100%'>$Name</td>
</tr>
<tr>
<td class='first'>EDIPI Number:</td>
<td>$Pre</td>
</tr>
<tr class='alt'>
<td class='first'>Base Name:</td>
<td>$Base</td>
</tr>
<tr>
<td class='first'>Email Address:</td>
<td><A HREF=mailto:$Email>$Email</A></td>
</tr>
<tr class='alt'>
<td class='first'>Mail Category:</td>
<td>$Cat</td>
</tr>
<tr>
<td class='first'>Box Size Limit:</td>
<td>$MailSize MB</td>
</tr>
<tr class='alt'>
<td class='first'>Account Locked:</td>
<td>$Locked</td>
</tr>
<tr>
<td class='first'>Account Enabled:</td>
<td>$Enabled</td>
</tr>
<tr class='alt'>
<td class='first'>Office Phone:</td>
<td>$Number</td>
</tr>
</table>
<a name='installedPrograms'><h1>Installed Programs</h1></a>
<table border='0' cellspacing='0' align='center' Width=100%>
<tr class='alt' Width=100%>
<td class='first'>
<DT><CENTER>Display Name</CENTER></DT>
</td>
<td class='first'>
<DT><CENTER>Display Version</CENTER></DT>
</td>
<td class='first'>
<DT><CENTER>Install Date</CENTER></DT>
</td>
<td class='first'>
<DT><CENTER>Publisher</CENTER></DT>
</td>
<td class='first'>
<DT><CENTER>Help Link</CENTER></DT>
</td>
</tr>
"+$ProgHTML

ConvertTo-Html -Head $Head -Body $Body | Out-File C:\Users\$LocalEDI\Desktop\$SysName.html

Start-Process -FilePath C:\Users\$LocalEDI\Desktop\$SysName.html