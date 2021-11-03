#=====Created by TSgt Simonds and SSgt Pelletier==================================
#Purpose:  multi Function Script for "almost all" of your 
#CST needs, with some additional harmless shenanigans.
#Note:  use this wisely and don't break anything, please.  
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
    ElseIf ($Comp -eq "Worley"){$Compname = "XLWUW-491S3B"}
    

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
    ElseIf ($User.UserName -eq "AREA52\1473682512N"){$Name = "Worley"}
    

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
Function EnableBIOSComponent
{
$Device=Get-WmiObject -computername $Computer -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface
$Device.SetBIOSSetting("$S","$EorD")
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

$Computers = Get-Content "C:\users\1383807847.adm\desktop\scripts\computer.txt"

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
Function CDROM
{
Invoke-Command -ComputerName $Computer -ScriptBlock {

$sh = New-Object -ComObject "Shell.Application"
$sh.Namespace(17).Items() | 
    Where-Object { $_.Type -eq "CD Drive" } | 
        foreach { $_.InvokeVerb("Eject") }
 }
 }
#===================================================================
Function CDROM2
{
$Computer = Get-Content C:\users\1252862141.adm\Desktop\Scripts1\Pop.txt

Invoke-Command -ComputerName $Computer -ScriptBlock {

$items = (New-Object -com "WMPlayer.OCX.7").cdromcollection.item(0)            
$items.eject()  
}
}
#===================================================================
Function RoboCopy
{
<#
Description:
Leverages the command-line utility Robocopy.
Includes the ability to copy file attributes along with the NTFS permissions, to mirror the content of an entire folder hierarchy across local volumes
or over a network excluding certain file types, copying files above or below a certain age or size, monitoring the source for changes, giving detailed
report with an option to output both to console window and log file.

Features:
- supports spaces in the file name
- select and copy text from the content output box
- recommended options
- advanced options
- enable/disable file logging
- generates log file name (current date + source folder name)
- opens the current job logfile in text editor
- parses the current log file and shows only ERROR messages

1.0.1 Updates:
- save preferences function
- progressbar
- stop Robocopy function

Version: 1.0.1 - 1/22/2015
Author: Nikolay Petkov
Blog: http://power-shell.com/
Link: http://power-shell.com/2014/powershell-gui-tools/robocopy-gui-tool/

License Info:
Copyright (c) power-shell.com 2014.
Distributed under the MIT License (http://opensource.org/licenses/MIT)
#>
        Write-Host " "
        Write-Host -ForegroundColor Gray "The following describes your 'Copy Options':"
        Write-Host " "
        Write-Host "/S - Copies subdirectories. Note that this option excludes empty directories."
        Write-Host "/E - Copies subdirectories. Note that this option includes empty directories."
        Write-Host "/B - Copies files in Backup mode."
        Write-Host "/SEC - Copies files with security (equivalent to /copy:DATS)."
        Write-Host "/COPYALL - Copies all file information (equivalent to /copy:DATSOU)."
        Write-Host "/NOCOPY - Copies no file information (useful with /purge)."
        Write-Host "/SECFIX - Fixes file security on all files, even skipped ones."
        Write-Host "/PURGE - Deletes destination files and directories that no longer exist in the source."
        Write-Host "/MIR - Mirrors a directory tree (equivalent to /e plus /purge)."
        Write-Host "/MOV - Moves files, and deletes them from the source after they are copied."
        Write-Host "/MOVE - Moves files and directories, and deletes them from the source after they are copied."
        Write-Host "/MT:8 - Creates multi-threaded copies with N threads. N must be an integer between 1 and 128. The default value for N is 8."
        Write-Host " "
        Write-Host -ForegroundColor Gray "The following describes your 'File Selection Options':"
        Write-Host " "
        Write-Host "/A - Copies only files for which the Archive attribute is set."
        Write-Host "/M - Copies only files for which the Archive attribute is set, and resets the Archive attribute."
        Write-Host "/XC - Excludes changed files."
        Write-Host "/XN - Excludes newer files."
        Write-Host "/XO - Excludes older files."
        Write-Host "/XX - Excludes extra files and directories."
        Write-Host "/XL - Excludes "lonely" files and directories."
        Write-Host "/IS - Includes the same files."
        Write-Host "/IT - Includes "tweaked" files."
        Write-Host "/XJ - Excludes junction points, which are normally included by default."
        Write-Host "/XJD - Excludes junction points for directories."
        Write-Host "/XJF - Excludes junction points for files."
        Write-Host " " 
        Write-Host -ForegroundColor Gray "Advanced Options: only use if you're familiar with RoboCopy commands."




[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(850,520)
$Form.Text = "PowerCopy (v1.0.1)"
$Form.StartPosition = "CenterScreen" #loads the window in the center of the screen
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
#$Image = [system.drawing.image]::FromFile("\\\")
#$Form.BackgroundImage = $Image
#$Form.BackgroundImageLayout = "Zoom"    # Options: None, Tile, Center, Stretch, Zoom
$Form.MinimizeBox = $False
$Form.MaximizeBox = $False
$Form.WindowState = "Normal"    # Options: Maximized, Minimized, Normal
$Form.SizeGripStyle = "Hide"    # Options: Auto, Hide, Show
$Form.Icon = $Icon
$Form.BackColor = "#CCCCCC"
#$Form.Opacity = 0.7
#$Font = New-Object System.Drawing.Font("Times New Roman",24,[System.Drawing.FontStyle]::Italic)    # Options: Regular, Bold, Italic, Underline, Strikeout
#$Form.Font = $Font


#Start Robocopy function
function robocopy {
begin {
#Recommended options
if ($checkboxNP.Checked) {$switchNP = "/NP"} else {$switchNP = $null} #/NP :: No Progress - don't display percentage copied


#Copy options
if ($checkboxS.Checked) {$switchS = "/S"} else {$switchS = $null} #/S :: copy Subdirectories, but not empty ones
if ($checkboxE.Checked) {$switchE = "/E"} else {$switchE = $null} #/E :: copy subdirectories, including empty ones. /E is including /S
if ($checkboxB.Checked) {$switchB = "/B"} else {$switchB = $null} #/B :: copy files in Backup mode
if ($checkboxSEC.Checked) {$switchSEC = "/SEC"} else {$switchSEC = $null} #/SEC :: copy files with SECurity (equivalent to /COPY:DATS)
if ($checkboxCOPYALL.Checked) {$switchCOPYALL = "/COPYALL"} else {$switchCOPYALL = $null} #COPY ALL file info (equivalent to /COPY:DATSOU)
if ($checkboxNOCOPY.Checked) {$switchNOCOPY = "/NOCOPY"} else {$switchNOCOPY = $null} #COPY NO file info (useful with /PURGE)
if ($checkboxSECFIX.Checked) {$switchSECFIX = "/SECFIX"} else {$switchSECFIX = $null} #FIX file SECurity on all files, even skipped files
if ($checkboxPURGE.Checked) {$switchPURGE = "/PURGE"} else {$switchPURGE = $null} #delete dest files/dirs that no longer exist in source
if ($checkboxMIR.Checked) {$switchMIR = "/MIR"} else {$switchMIR = $null} #MIRror a directory tree (equivalent to /E plus /PURGE)
if ($checkboxMOV.Checked) {$switchMOV = "/MOV"} else {$switchMOV = $null} #MOVe files (delete from source after copying)
if ($checkboxMOVE.Checked) {$switchMOVE = "/MOVE"} else {$switchMOVE = $null} #MOVE files AND dirs (delete from source after copying)
if ($checkboxMT.Checked) {$switchMT = "/MT"} else {$switchMT = $null} #Do multi-threaded copies with n threads (default 8)
if ($checkboxA.Checked) {$switchA = "/A"} else {$switchA = $null} #copy only files with the Archive attribute set
if ($checkboxM.Checked) {$switchM = "/M"} else {$switchM = $null} #copy only files with the Archive attribute and reset it
if ($checkboxXC.Checked) {$switchXC = "/XC"} else {$switchXC = $null} #eXclude Changed files
if ($checkboxXN.Checked) {$switchXN = "/XN"} else {$switchXN = $null} #eXclude Newer files
if ($checkboxXO.Checked) {$switchXO = "/XO"} else {$switchXO = $null} #eXclude Older files
if ($checkboxXX.Checked) {$switchXX = "/XX"} else {$switchXX = $null} #eXclude eXtra files and directories
if ($checkboxXL.Checked) {$switchXL = "/XL"} else {$switchXL = $null} #eXclude Lonely files and directories
if ($checkboxIS.Checked) {$switchIS = "/IS"} else {$switchIS = $null} #Include Same files
if ($checkboxIT.Checked) {$switchIT = "/IT"} else {$switchIT = $null} #Include Tweaked files
if ($checkboxXJ.Checked) {$switchXJ = "/XJ"} else {$switchXJ = $null} # eXclude Junction points. (normally included by default)
if ($checkboxXJD.Checked) {$switchXJD = "/XJD"} else {$switchXJD = $null} #eXclude Junction points for Directories
if ($checkboxXJF.Checked) {$switchXJF = "/XJF"} else {$switchXJF = $null} #eXclude Junction points for Files
if ($checkboxL.Checked) {$switchL = "/L"} else {$switchL = $null} #List only - don't copy, timestamp or delete any files
if ($checkboxX.Checked) {$switchX = "/X"} else {$switchX = $null} #report all eXtra files, not just those selected
if ($checkboxV.Checked) {$switchV = "/V"} else {$switchV = $null} #produce Verbose output, showing skipped files
if ($checkboxTS.Checked) {$switchTS = "/TS"} else {$switchTS = $null} #include source file Time Stamps in the output
if ($checkboxFP.Checked) {$switchFP = "/FP"} else {$switchFP = $null} #include Full Pathname of files in the output
if ($checkboxBYTES.Checked) {$switchBYTES = "/BYTES"} else {$switchBYTES = $null} #Print sizes as bytes
if ($checkboxR.Checked) {$switchR = "/R:3"} else {$switchR = $null} #number of Retries on failed copies: default 1 million
if ($checkboxW.Checked) {$switchW = "/W:1"} else {$switchW = $null} #Wait time between retries: default is 30 seconds

#Additional options
if ($InputAdvancedOptions.Text) {$switchAddition = $InputAdvancedOptions.Text.split(' ')} else {$switchAddition = $null}

#Log File Function
if (($checkboxLog.Checked -and $InputLogFile.Text))
{
if(!(Test-Path -Path $InputLogFile.Text)){
$checkpath ="`nError: The logfile path " + """" + $InputLogFile.Text + """" + " doesn't exist!`n"
}
$logfile = $InputLogFile.Text + "\" + ((Get-Date).ToString('yyyy-MM-dd')) + "_" + $InputSource.Text.Split('\')[-1].Replace(" ","_") + ".txt"
$switchlogfile = "/TEE", "/LOG+:$logfile"
}
else {$switchlogfile = $null}
if (!($logfile)) {$checklog = "  Log File : The logging is not enabled."
}
$outputBox.text = $checklog, $checkpath}
process {
#count the source files
$outputBox.text = " Preparing Robocopy. Please wait..."
if ($InputSource.Text -notlike $null) {
$sourcefiles=robocopy.exe $InputSource.Text $InputSource.Text /L /S /NJH /BYTES /FP /NC /NDL /TS /XJ /R:0 /W:0
If ($sourcefiles[-5] -match '^\s{3}Files\s:\s+(?<Count>\d+).*') {$filecount=$matches.Count}
}
$outputBox.Focus()
$run = robocopy.exe $InputSource.Text $InputTarget.Text $switchNP $switchR $switchW $switchS $switchE $switchB $switchSEC $switchCOPYALL $switchNOCOPY `
$switchSECFIX $switchPURGE $switchMIR $switchMOV $switchMOVE $switchMT $switchA $switchM $switchXC $switchXN $switchXO $switchXX `
$switchXL $switchIS $switchIT $switchXJ $switchXJD $switchXJF $switchL $switchX $switchV $switchTS $switchFP $switchBYTES $switchAddition $switchLogfile | foreach {
$ErrorActionPreference = "silentlycontinue"
#calculate percentage
$i++
[int]$pct = ($i/$filecount)*100
#update the progress bar
$progressbar.Value = ($pct)
$outputBox.AppendText($_ + "`r`n")
[void] [System.Windows.Forms.Application]::DoEvents()
}
}
end {$progressbar.Value = 100}
} #end robocopy function
               
#Robocopy Help function
function robocopyhelp {
$help = robocopy.exe /?
$outputBox.text = $help |Out-String
}
#Open log function
function openlog {
$logfile = $InputLogFile.Text + "\" + ((Get-Date).ToString('yyyy-MM-dd')) + "_" + $InputSource.Text.Split('\')[-1].Replace(" ","_") + ".txt"
if(!(Test-Path $logfile)){$outputBox.text = "There is no logfile for the current job."}
else
{$openlog = notepad.exe $logfile}
}
#Show Errors function
function showerrors {
$logfile = $InputLogFile.Text + "\" + ((Get-Date).ToString('yyyy-MM-dd')) + "_" + $InputSource.Text.Split('\')[-1].Replace(" ","_") + ".txt"
if
(!(Test-Path $logfile)) {$outputBox.text = "There is no logfile for the current job."}
else
{$logcontent = Get-Content $logfile
if ($errors = $logcontent | Select-String -Pattern "ERROR " -Context 0,1 |Out-String) {$outputBox.text = $errors}
else {$outputBox.text = "No errors found."}
}
}
#Stop Robocopy function
function stoprobocopy {
if (get-process -Name robocopy -ErrorAction SilentlyContinue) {Stop-Process -Name robocopy -Force
$timestamp = (Get-Date).ToString('yyyy/MM/dd hh:mm:ss')
$outputBox.AppendText("`n`r$timestamp Robocopy process has been terminated.")}
if ($logfile) {
Add-Content $logfile "`n`r$timestamp ERROR Robocopy process has been terminated."}
} #end stop Robocopy function

#Save Options function
$Scriptpath = $myInvocation.InvocationName
Function saveoptions {
try {
$saveadvanced = """" + $InputAdvancedOptions.Text.ToString() + """"
$savelogpath = """" + $InputLogFile.Text.ToString() + """"
$noerror = $true
(Get-Content $Scriptpath) | ForEach-Object {
if ($_ | Select-String '^.checkboxS.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxS.Checked}
elseif ($_ | Select-String '^.checkboxE.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxE.Checked}
elseif ($_ | Select-String '^.checkboxB.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxB.Checked}
elseif ($_ | Select-String '^.checkboxSEC.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxSEC.Checked}
elseif ($_ | Select-String '^.checkboxCOPYALL.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxCOPYALL.Checked}
elseif ($_ | Select-String '^.checkboxNOCOPY.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxNOCOPY.Checked}
elseif ($_ | Select-String '^.checkboxSECFIX.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxSECFIX.Checked}
elseif ($_ | Select-String '^.checkboxPURGE.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxPURGE.Checked}
elseif ($_ | Select-String '^.checkboxMIR.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxMIR.Checked}
elseif ($_ | Select-String '^.checkboxMOV.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxMOV.Checked}
elseif ($_ | Select-String '^.checkboxMOVE.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxMOVE.Checked}
elseif ($_ | Select-String '^.checkboxMT.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxMT.Checked}
elseif ($_ | Select-String '^.checkboxA.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxA.Checked}
elseif ($_ | Select-String '^.checkboxM.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxM.Checked}
elseif ($_ | Select-String '^.checkboxXC.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxXC.Checked}
elseif ($_ | Select-String '^.checkboxXN.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxXN.Checked}
elseif ($_ | Select-String '^.checkboxXO.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxXO.Checked}
elseif ($_ | Select-String '^.checkboxXX.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxXX.Checked}
elseif ($_ | Select-String '^.checkboxXL.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxXL.Checked}
elseif ($_ | Select-String '^.checkboxIS.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxIS.Checked}
elseif ($_ | Select-String '^.checkboxIT.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxIT.Checked}
elseif ($_ | Select-String '^.checkboxXJ.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxXJ.Checked}
elseif ($_ | Select-String '^.checkboxXJD.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxXJD.Checked}
elseif ($_ | Select-String '^.checkboxXJF.Checked') {$_ -replace ($_ -split "=")[1].substring(1), $checkboxXJF.Checked}
elseif ($_ | Select-String '^.InputAdvancedOptions.Text') {$_.Replace($_.Split("=")[1], $saveadvanced)}
elseif ($_ | Select-String '^.InputLogFile.Text') {$_.Replace($_.Split("=")[1], $savelogpath)}
else {$_}


} | Set-Content $Scriptpath -erroraction stop
} catch {
[System.Windows.Forms.MessageBox]::Show("An error occurred while saving your preferences.","Save Preferences", "Ok", "Error")
$noerror = $false
        }
if ($noerror) {
[System.Windows.Forms.MessageBox]::Show("Your preferences have been saved.","Save Preferences", "Ok", "Information")
              }
}#end Save Options function

#checkbox group boxes

#copy options group box
$copyGroupBox = New-Object System.Windows.Forms.GroupBox
$copyGroupBox.Location = New-Object System.Drawing.Size(210,15) 
$copyGroupBox.size = New-Object System.Drawing.Size(220,110) 
$copyGroupBox.text = "Copy Options" 
$Form.Controls.Add($copyGroupBox)

#file selection options group box
$FileSelectionGroupBox = New-Object System.Windows.Forms.GroupBox
$FileSelectionGroupBox.Location = New-Object System.Drawing.Size(440,15) 
$FileSelectionGroupBox.size = New-Object System.Drawing.Size(185,110) 
$FileSelectionGroupBox.text = "File Selection Options" 
$Form.Controls.Add($FileSelectionGroupBox)

#recommended options group box
$RecommendedGroupBox = New-Object System.Windows.Forms.GroupBox
$RecommendedGroupBox.Location = New-Object System.Drawing.Size(640,15)
$RecommendedGroupBox.size = New-Object System.Drawing.Size(190,50)
$RecommendedGroupBox.text = "Recommended Options" 
$Form.Controls.Add($RecommendedGroupBox)

#advanced options groupBox
$AdvancedGroupBox = New-Object System.Windows.Forms.GroupBox
$AdvancedGroupBox.Location = New-Object System.Drawing.Size(640,75)
$AdvancedGroupBox.Size = New-Object System.Drawing.Size(190,50)
$AdvancedGroupBox.Text = "Advanced Options:" 
$Form.Controls.Add($AdvancedGroupBox)

#advanced options input
$InputAdvancedOptions = New-Object System.Windows.Forms.TextBox
$InputAdvancedOptions.Text=""
$InputAdvancedOptions.Location = New-Object System.Drawing.Size(10,20) 
$InputAdvancedOptions.Size = New-Object System.Drawing.Size(170,30) 
$AdvancedGroupBox.Controls.Add($InputAdvancedOptions)

#log file path groupbox
$LogFileGroupbox = New-Object System.Windows.Forms.GroupBox
$LogFileGroupbox.Text="Logfile Path"
$LogFileGroupbox.Location = New-Object System.Drawing.Size(640,170) 
$LogFileGroupbox.Size = New-Object System.Drawing.Size(190,50) 
$Form.Controls.Add($LogFileGroupbox)

#log file path input
$InputLogFile = New-Object System.Windows.Forms.TextBox
$InputLogFile.Text="c:\users\1252862141.adm\desktop"
$InputLogFile.Location = New-Object System.Drawing.Size(10,20) 
$InputLogFile.Size = New-Object System.Drawing.Size(170,30) 
$LogFileGroupbox.Controls.Add($InputLogFile)

#logging options group box
$LoggingGroupBox = New-Object System.Windows.Forms.GroupBox
$LoggingGroupBox.Location = New-Object System.Drawing.Size(640,230)
$LoggingGroupBox.size = New-Object System.Drawing.Size(190,70)
$LoggingGroupBox.text = "Logging Options" 
$Form.Controls.Add($LoggingGroupBox)
#end group boxes

#check boxes

#Robocopy options check boxes

#start copy options
$checkboxS = New-Object System.Windows.Forms.checkbox
$checkboxS.Location = New-Object System.Drawing.Size(10,20)
$checkboxS.Size = New-Object System.Drawing.Size(50,20)
$checkboxS.Checked=$False
$checkboxS.Text = "/S"
$copyGroupBox.Controls.Add($checkboxS)

$checkboxE = New-Object System.Windows.Forms.checkbox
$checkboxE.Location = New-Object System.Drawing.Size(10,40)
$checkboxE.Size = New-Object System.Drawing.Size(50,20)
$checkboxE.Checked=$True
$checkboxE.Text = "/E"
$copyGroupBox.Controls.Add($checkboxE)

$checkboxB = New-Object System.Windows.Forms.checkbox
$checkboxB.Location = New-Object System.Drawing.Size(10,60)
$checkboxB.Size = New-Object System.Drawing.Size(50,20)
$checkboxB.Checked=$True
$checkboxB.Text = "/B"
$copyGroupBox.Controls.Add($checkboxB)

$checkboxSEC = New-Object System.Windows.Forms.checkbox
$checkboxSEC.Location = New-Object System.Drawing.Size(10,80)
$checkboxSEC.Size = New-Object System.Drawing.Size(50,20)
$checkboxSEC.Checked=$True
$checkboxSEC.Text = "/SEC"
$copyGroupBox.Controls.Add($checkboxSEC)

#COPY ALL file info (equivalent to /COPY:DATSOU)
$checkboxCOPYALL = New-Object System.Windows.Forms.checkbox
$checkboxCOPYALL.Location = New-Object System.Drawing.Size(70,20)
$checkboxCOPYALL.Size = New-Object System.Drawing.Size(80,20)
$checkboxCOPYALL.Checked=$False
$checkboxCOPYALL.Text = "/COPYALL"
$copyGroupBox.Controls.Add($checkboxCOPYALL)

#COPY NO file info (useful with /PURGE)
$checkboxNOCOPY = New-Object System.Windows.Forms.checkbox
$checkboxNOCOPY.Location = New-Object System.Drawing.Size(70,40)
$checkboxNOCOPY.Size = New-Object System.Drawing.Size(80,20)
$checkboxNOCOPY.Checked=$False
$checkboxNOCOPY.Text = "/NOCOPY"
$copyGroupBox.Controls.Add($checkboxNOCOPY)

#FIX file SECurity on all files, even skipped files
$checkboxSECFIX = New-Object System.Windows.Forms.checkbox
$checkboxSECFIX.Location = New-Object System.Drawing.Size(70,60)
$checkboxSECFIX.Size = New-Object System.Drawing.Size(80,20)
$checkboxSECFIX.Checked=$True
$checkboxSECFIX.Text = "/SECFIX"
$copyGroupBox.Controls.Add($checkboxSECFIX)

#delete dest files/dirs that no longer exist in source
$checkboxPURGE = New-Object System.Windows.Forms.checkbox
$checkboxPURGE.Location = New-Object System.Drawing.Size(70,80)
$checkboxPURGE.Size = New-Object System.Drawing.Size(80,20)
$checkboxPURGE.Checked=$False
$checkboxPURGE.Text = "/PURGE"
$copyGroupBox.Controls.Add($checkboxPURGE)

#MIRror a directory tree (equivalent to /E plus /PURGE)
$checkboxMIR = New-Object System.Windows.Forms.checkbox
$checkboxMIR.Location = New-Object System.Drawing.Size(157,20)
$checkboxMIR.Size = New-Object System.Drawing.Size(60,20)
$checkboxMIR.Checked=$True
$checkboxMIR.Text = "/MIR"
$copyGroupBox.Controls.Add($checkboxMIR)

#MOVE files (delete from source after copying)
$checkboxMOV = New-Object System.Windows.Forms.checkbox
$checkboxMOV.Location = New-Object System.Drawing.Size(157,40)
$checkboxMOV.Size = New-Object System.Drawing.Size(60,20)
$checkboxMOV.Checked=$False
$checkboxMOV.Text = "/MOV"
$copyGroupBox.Controls.Add($checkboxMOV)

#MOVE files AND dirs (delete from source after copying)
$checkboxMOVE = New-Object System.Windows.Forms.checkbox
$checkboxMOVE.Location = New-Object System.Drawing.Size(157,60)
$checkboxMOVE.Size = New-Object System.Drawing.Size(60,20)
$checkboxMOVE.Checked=$False
$checkboxMOVE.Text = "/MOVE"
$copyGroupBox.Controls.Add($checkboxMOVE)

#Do multi-threaded copies with n threads (default 8)
$checkboxMT = New-Object System.Windows.Forms.checkbox
$checkboxMT.Location = New-Object System.Drawing.Size(157,80)
$checkboxMT.Size = New-Object System.Drawing.Size(60,20)
$checkboxMT.Checked=$True
$checkboxMT.Text = "/MT:8"
$copyGroupBox.Controls.Add($checkboxMT)

#end copy options

#start file selection options check boxes

#copy only files with the Archive attribute set
$checkboxA = New-Object System.Windows.Forms.checkbox
$checkboxA.Location = New-Object System.Drawing.Size(10,20)
$checkboxA.Size = New-Object System.Drawing.Size(50,20)
$checkboxA.Checked=$False
$checkboxA.Text = "/A"
$FileSelectionGroupBox.Controls.Add($checkboxA)

#copy only files with the Archive attribute and reset it
$checkboxM = New-Object System.Windows.Forms.checkbox
$checkboxM.Location = New-Object System.Drawing.Size(10,40)
$checkboxM.Size = New-Object System.Drawing.Size(50,20)
$checkboxM.Checked=$False
$checkboxM.Text = "/M"
$FileSelectionGroupBox.Controls.Add($checkboxM)

#eXclude changed files
$checkboxXC = New-Object System.Windows.Forms.checkbox
$checkboxXC.Location = New-Object System.Drawing.Size(10,60)
$checkboxXC.Size = New-Object System.Drawing.Size(50,20)
$checkboxXC.Checked=$False
$checkboxXC.Text = "/XC"
$FileSelectionGroupBox.Controls.Add($checkboxXC)

#eXclude Newer files
$checkboxXN = New-Object System.Windows.Forms.checkbox
$checkboxXN.Location = New-Object System.Drawing.Size(10,80)
$checkboxXN.Size = New-Object System.Drawing.Size(50,20)
$checkboxXN.Checked=$False
$checkboxXN.Text = "/XN"
$FileSelectionGroupBox.Controls.Add($checkboxXN)

#eXclude Older files
$checkboxXO = New-Object System.Windows.Forms.checkbox
$checkboxXO.Location = New-Object System.Drawing.Size(70,20)
$checkboxXO.Size = New-Object System.Drawing.Size(50,20)
$checkboxXO.Checked=$False
$checkboxXO.Text = "/XO"
$FileSelectionGroupBox.Controls.Add($checkboxXO)

#eXclude eXtra files and directories
$checkboxXX = New-Object System.Windows.Forms.checkbox
$checkboxXX.Location = New-Object System.Drawing.Size(70,40)
$checkboxXX.Size = New-Object System.Drawing.Size(50,20)
$checkboxXX.Checked=$False
$checkboxXX.Text = "/XX"
$FileSelectionGroupBox.Controls.Add($checkboxXX)

#eXclude Lonely files and directories
$checkboxXL = New-Object System.Windows.Forms.checkbox
$checkboxXL.Location = New-Object System.Drawing.Size(70,60)
$checkboxXL.Size = New-Object System.Drawing.Size(50,20)
$checkboxXL.Checked=$False
$checkboxXL.Text = "/XL"
$FileSelectionGroupBox.Controls.Add($checkboxXL)

#Include Same files
$checkboxIS = New-Object System.Windows.Forms.checkbox
$checkboxIS.Location = New-Object System.Drawing.Size(70,80)
$checkboxIS.Size = New-Object System.Drawing.Size(50,20)
$checkboxIS.Checked=$False
$checkboxIS.Text = "/IS"
$FileSelectionGroupBox.Controls.Add($checkboxIS)

#Include Tweaked files
$checkboxIT = New-Object System.Windows.Forms.checkbox
$checkboxIT.Location = New-Object System.Drawing.Size(130,20)
$checkboxIT.Size = New-Object System.Drawing.Size(50,20)
$checkboxIT.Checked=$False
$checkboxIT.Text = "/IT"
$FileSelectionGroupBox.Controls.Add($checkboxIT)

#eXclude Junction points
$checkboxXJ = New-Object System.Windows.Forms.checkbox
$checkboxXJ.Location = New-Object System.Drawing.Size(130,40)
$checkboxXJ.Size = New-Object System.Drawing.Size(50,20)
$checkboxXJ.Checked=$False
$checkboxXJ.Text = "/XJ"
$FileSelectionGroupBox.Controls.Add($checkboxXJ)

#eXclude Junction points for Directories
$checkboxXJD = New-Object System.Windows.Forms.checkbox
$checkboxXJD.Location = New-Object System.Drawing.Size(130,60)
$checkboxXJD.Size = New-Object System.Drawing.Size(50,20)
$checkboxXJD.Checked=$False
$checkboxXJD.Text = "/XJD"
$FileSelectionGroupBox.Controls.Add($checkboxXJD)

#eXclude Junction points for Files
$checkboxXJF = New-Object System.Windows.Forms.checkbox
$checkboxXJF.Location = New-Object System.Drawing.Size(130,80)
$checkboxXJF.Size = New-Object System.Drawing.Size(50,20)
$checkboxXJF.Checked=$False
$checkboxXJF.Text = "/XJF"
$FileSelectionGroupBox.Controls.Add($checkboxXJF)

#end Robocopy file selection options

#start logging options

#Enable Logging checkbox
$checkboxLog = New-Object System.Windows.Forms.checkbox
$checkboxLog.Location = New-Object System.Drawing.Size(640,140)
$checkboxLog.Size = New-Object System.Drawing.Size(110,20)
$checkboxLog.Checked=$True
$checkboxLog.Text = "Enable Logging"
$Form.Controls.Add($checkboxLog)

#List only - don't copy, timestamp or delete any files
$checkboxL = New-Object System.Windows.Forms.checkbox
$checkboxL.Location = New-Object System.Drawing.Size(10,20)
$checkboxL.Size = New-Object System.Drawing.Size(50,20)
$checkboxL.Checked=$False
$checkboxL.Text = "/L"
$LoggingGroupBox.Controls.Add($checkboxL)

#report all eXtra files, not just those selected
$checkboxX = New-Object System.Windows.Forms.checkbox
$checkboxX.Location = New-Object System.Drawing.Size(10,40)
$checkboxX.Size = New-Object System.Drawing.Size(50,20)
$checkboxX.Checked=$False
$checkboxX.Text = "/X"
$LoggingGroupBox.Controls.Add($checkboxX)

#produce Verbose output, showing skipped files
$checkboxV = New-Object System.Windows.Forms.checkbox
$checkboxV.Location = New-Object System.Drawing.Size(70,20)
$checkboxV.Size = New-Object System.Drawing.Size(50,20)
$checkboxV.Checked=$False
$checkboxV.Text = "/V"
$LoggingGroupBox.Controls.Add($checkboxV)

#include source file Time Stamps in the output
$checkboxTS = New-Object System.Windows.Forms.checkbox
$checkboxTS.Location = New-Object System.Drawing.Size(70,40)
$checkboxTS.Size = New-Object System.Drawing.Size(50,20)
$checkboxTS.Checked=$False
$checkboxTS.Text = "/TS"
$LoggingGroupBox.Controls.Add($checkboxTS)

#include Full Pathname of files in the output
$checkboxFP = New-Object System.Windows.Forms.checkbox
$checkboxFP.Location = New-Object System.Drawing.Size(125,20)
$checkboxFP.Size = New-Object System.Drawing.Size(50,20)
$checkboxFP.Checked=$False
$checkboxFP.Text = "/FP"
$LoggingGroupBox.Controls.Add($checkboxFP)

#Print sizes as bytes
$checkboxBYTES = New-Object System.Windows.Forms.checkbox
$checkboxBYTES.Location = New-Object System.Drawing.Size(125,40)
$checkboxBYTES.Size = New-Object System.Drawing.Size(63,20)
$checkboxBYTES.Checked=$False
$checkboxBYTES.Text = "/BYTES"
$LoggingGroupBox.Controls.Add($checkboxBYTES)

#end logging options

#start recommended options
#No Progress - don't display percentage copied
$checkboxNP = New-Object System.Windows.Forms.checkbox
$checkboxNP.Location = New-Object System.Drawing.Size(10,20)
$checkboxNP.Size = New-Object System.Drawing.Size(50,20)
$checkboxNP.Checked=$True
$checkboxNP.Text = "/NP"
$RecommendedGroupBox.Controls.Add($checkboxNP)

#start recommended options

#number of Retries on failed copies: default 1 million
$checkboxR = New-Object System.Windows.Forms.checkbox
$checkboxR.Location = New-Object System.Drawing.Size(70,20)
$checkboxR.Size = New-Object System.Drawing.Size(50,20)
$checkboxR.Checked=$True
$checkboxR.Text = "/R:3"
$RecommendedGroupBox.Controls.Add($checkboxR)

#number of Retries on failed copies: default 1 million
$checkboxW = New-Object System.Windows.Forms.checkbox
$checkboxW.Location = New-Object System.Drawing.Size(130,20)
$checkboxW.Size = New-Object System.Drawing.Size(55,20)
$checkboxW.Checked=$True
$checkboxW.Text = "/W:1"
$RecommendedGroupBox.Controls.Add($checkboxW)

#end recommended options

#Text fields

#Source path label
$InputSourceLabel = New-Object System.Windows.Forms.Label
$InputSourceLabel.Text="Source Path:"
$InputSourceLabel.Location = New-Object System.Drawing.Size(15,15) 
$InputSourceLabel.Size = New-Object System.Drawing.Size(170,15) 
$Form.Controls.Add($InputSourceLabel)

#Source path input
$InputSource = New-Object System.Windows.Forms.TextBox
$InputSource.Text=""
$InputSource.Location = New-Object System.Drawing.Size(15,30) 
$InputSource.Size = New-Object System.Drawing.Size(180,20) 
$Form.Controls.Add($InputSource)

#Target path label
$InputTargetLabel = New-Object System.Windows.Forms.Label
$InputTargetLabel.Text="Destination Path:"
$InputTargetLabel.Location = New-Object System.Drawing.Size(15,55) 
$InputTargetLabel.Size = New-Object System.Drawing.Size(170,15) 
$Form.Controls.Add($InputTargetLabel)

#Target path input
$InputTarget = New-Object System.Windows.Forms.TextBox
$InputTarget.Text=""
$InputTarget.Location = New-Object System.Drawing.Size(15,70) 
$InputTarget.Size = New-Object System.Drawing.Size(180,30) 
$Form.Controls.Add($InputTarget)

#Output box
$outputBox = New-Object System.Windows.Forms.RichTextBox 
$outputBox.Location = New-Object System.Drawing.Size(15,150) 
$outputBox.Size = New-Object System.Drawing.Size(610,290)
$outputBox.MultiLine = $True
#$outputBox.WordWrap = $False
$outputBox.ScrollBars = "Both"
$outputBox.Font = "Courier New"
$Form.Controls.Add($outputBox)

########### HomePage URL Label
$URLLabel = New-Object System.Windows.Forms.LinkLabel 
$URLLabel.Location = New-Object System.Drawing.Size(735,455) 
$URLLabel.Size = New-Object System.Drawing.Size(200,30)
$URLLabel.LinkColor = "#000000" 
$URLLabel.ActiveLinkColor = "Blue"
$URLLabel.Text = "Check for updates" 
$URLLabel.add_Click({[system.Diagnostics.Process]::start("http:\\power-shell.com")}) 
$Form.Controls.Add($URLLabel) 

#end text fields

#Start buttons

#Button Start Robocopy
$ButtonStart = New-Object System.Windows.Forms.Button 
$ButtonStart.Location = New-Object System.Drawing.Size(640,360) 
$ButtonStart.Size = New-Object System.Drawing.Size(190,80) 
#$ButtonStart.BackColor = "Green"
$ButtonStart.Text = "START ROBOCOPY" 
$ButtonStart.Add_Click({robocopy})
$Form.Controls.Add($ButtonStart) 

#Button Show Robocopy Help
$ButtonHelp = New-Object System.Windows.Forms.Button 
$ButtonHelp.Location = New-Object System.Drawing.Size(15,100) 
$ButtonHelp.Size = New-Object System.Drawing.Size(180,25) 
$ButtonHelp.Text = "Show Robocopy Help" 
$ButtonHelp.Add_Click({robocopyhelp})
$Form.Controls.Add($ButtonHelp)

#Button Save Robocopy Options
$ButtonSave = New-Object System.Windows.Forms.Button 
$ButtonSave.Location = New-Object System.Drawing.Size(640,310) 
$ButtonSave.Size = New-Object System.Drawing.Size(190,30) 
$ButtonSave.Text = "Save Options" 
$ButtonSave.Add_Click({saveoptions})
$Form.Controls.Add($ButtonSave) 

#Button Open Log
$ButtonOpenLog = New-Object System.Windows.Forms.Button 
$ButtonOpenLog.Location = New-Object System.Drawing.Size(15,450) 
$ButtonOpenLog.Size = New-Object System.Drawing.Size(110,25) 
$ButtonOpenLog.Text = "Open Logfile" 
$ButtonOpenLog.Add_Click({openlog})
$Form.Controls.Add($ButtonOpenLog)

#Button Show Errors
$ButtonErrors = New-Object System.Windows.Forms.Button 
$ButtonErrors.Location = New-Object System.Drawing.Size(140,450) 
$ButtonErrors.Size = New-Object System.Drawing.Size(110,25) 
$ButtonErrors.Text = "Show Errors" 
$ButtonErrors.Add_Click({showerrors})
$Form.Controls.Add($ButtonErrors)

#Button Stop Robocopy
$ButtonStop = New-Object System.Windows.Forms.Button 
$ButtonStop.Location = New-Object System.Drawing.Size(515,450) 
$ButtonStop.Size = New-Object System.Drawing.Size(110,25) 
#$ButtonStop.BackColor = "Red"
$ButtonStop.Text = "Stop Robocopy" 
$ButtonStop.Add_Click({stoprobocopy})
$Form.Controls.Add($ButtonStop)

#end buttons

#start progres bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Name = 'ProgressBar'
$progressBar.Value = 0
$progressBar.Style="Continuous"
$progressBar.Location = New-Object System.Drawing.Size(270,450) 
$progressBar.Size = New-Object System.Drawing.Size(225,25)
#initialize a counter
$i=0
$Form.Controls.Add($progressBar)

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
}
#===================================================================
Function nslookup
{
Test-Connection -ComputerName "$Computer” -Count 3 -Delay 2 -TTL 255 -BufferSize 256 -ThrottleLimit 32 
}
#===================================================================
Function C$
{
invoke-item \\$Computer\C$
}
#===================================================================
Function GPUpdate
{
invoke-gpupdate -computer $Computer -randomdelayinminutes 0 -force 
}
#===================================================================
Function StopProcess
{
Invoke-Command -ComputerName $Computer -Script { param($service) stop-Process -name $service -force } -argumentlist $service
}
#===================================================================
Do
{
    Write-Host " "
    Write-Host "0  - Cls"
    Write-Host "1  - Send Message"
    Write-Host "2  - Missed Call"
    Write-Host "3  - Lock Computer"
    Write-Host "4  - Logoff User"
    Write-Host "5  - Reboot Computer"
    Write-Host "6  - Enable Local Admin"
    Write-Host "7  - Enable/Disable BIOS Component"
    Write-Host "8  - Enter-PSSession"
    Write-Host "9  - PC & Bldg"
    Write-Host "10 - PC & User"
    Write-Host "11 - CD-Rom"
    Write-Host "12 - Robo Copy"
    Write-Host "13 - Ping & NsLookup"
    Write-Host "14 - C$"
    Write-Host "15 - GPUpdate"
    Write-Host "16 - Stop-Process"
    Write-Host "17 - Exit"
    $Ans = Read-Host "Don't just sit there. Pick one"
    
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
        MissedCall
    }
    If ($Ans -eq 3)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        LockComputer
    }
    If ($Ans -eq 4)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        LogoffUser
    }
    If ($Ans -eq 5)
    {
        Write-Host
        $Computer = Read-Host "Computer"
        RebootComputer
    }
    If ($Ans -eq 6)
    {
        EnableLocalAdmin
    } 
    If ($Ans -eq 7)
    {
        Write-Host
        $SorM = Read-Host "Singe (s) or Multi (m) PCs?"
            If ($SorM -eq "s")
                {
                    $Computer = Read-Host "Computer"
                }
            If ($SorM -eq "m")
                {
                    $Multi = Read-Host "File Location"
                    $txt = Get-Content $multi
                    ForEach-Object ($Computer)
                } 
        $EnOrDis = Read-Host "Would you like to Enable (e) or Disable (d) a BIOS component?"
            
            If ($EnOrDis -eq "e")
                {
                    $EorD = "Enable"
                }
            
            If ($EnOrDis -eq "d")
                {
                    $EorD = "Disable"
                }
        
        Write-Host "1 - CD-ROM Boot"
        Write-Host "2 - Network (PXE) Boot"
        Write-Host "3 - NumLock on at boot"
        Write-Host "4 - Audio Device"
        Write-Host "5 - Integrated Microphone"
        Write-Host "6 - Internal Speakers"
        Write-Host "7 - Headphone Output"
        Write-Host "8 - Integrated Camera"
        Write-Host "9 - Fingerprint Device"
        Write-Host "10 - Prompt for Admin password on F9 (Boot Menu)"
        Write-Host "11 - Prompt for Admin password on F12 (Network Boot)" 
        $Which = Read-Host "Pick one"   
            If ($Which -eq 1)
                {
                    $S = "CD-ROM Boot"
                    EnableBIOSComponent
                 }
            If ($Which -eq 2)
                {
                    $S = "Network (PXE) Boot"
                    EnableBIOSComponent
                 }
            If ($Which -eq 3)
                {
                    $S = "NumLock on at boot"
                    EnableBIOSComponent
                 }
            If ($Which -eq 4)
                {
                    $S = "Audio Device"
                    EnableBIOSComponent
                 }
            If ($Which -eq 5)
                {
                    $S = "Integrated Microphone"
                    EnableBIOSComponent
                 }
            If ($Which -eq 6)
                {
                    $S = "Internal Speakers"
                    EnableBIOSComponent
                 }
            If ($Which -eq 7)
                {
                    $S = "Headphone Output"
                    EnableBIOSComponent
                 }
            If ($Which -eq 8)
                {
                    $S = "Integrated Camera"
                    EnableBIOSComponent
                 }
            If ($Which -eq 9)
                {
                    $S = "Fingerprint Device"
                    EnableBIOSComponent
                 }
             If ($Which -eq 10)
                {
                    $S = "Prompt for Admin password on F9 (Boot Menu)"
                    EnableBIOSComponent
                 }
             If ($Which -eq 11)
                {
                    $S = "Prompt for Admin password on F12 (Network Boot)"
                    EnableBIOSComponent
                 }
             
    }
    If ($Ans -eq 8)
    {
    Write-Host " "
    Write-Host "0 - Enter Computer Name"
    Write-Host "1 - Alex"
    Write-Host "2 - Arnold"
    Write-Host "3 - Ballentine"
    Write-Host "4 - Barnett"
    Write-Host "5 - Ben"
    Write-Host "6 - Brown"
    Write-Host "7 - Cain"
    Write-Host "8 - Dossa"
    Write-Host "9 - Ed"
    Write-Host "10 - Gail"
    Write-Host "11 - Ginger" 
    Write-Host "12 - Goldman" 
    Write-Host "13 - Grainger"
    Write-Host "14 - Cunningham"
    Write-Host "15 - Hiserodt"
    Write-Host "16 - Johns"
    Write-Host "17 - Lafond"
    Write-Host "18 - Linde"
    Write-Host "19 - Lozada"
    Write-Host "20 - Mac"
    Write-Host "21 - Mowry"
    Write-Host "22 - Oster"
    Write-Host "23 - Pelletier"
    Write-Host "24 - Pelletier Admin"
    Write-Host "25 - Scheffrin"
    Write-Host "26 - SMSgt Larry"
    Write-Host "27 - Thrift"
    Write-Host "28 - Walden"
    Write-Host "29 - Worley"
        $Pick = Read-Host "Choose a Victim/Customer"   
            If ($Pick -eq 1)
                {
                    $Computer = "XLWUW-491S6B"
                    Enter-PSSession
                 }
            If ($Pick -eq 2)
                {
                    $Computer = "XLWUL-4422Z4"
                    Enter-PSSession
                 }
            If ($Pick -eq 3)
                {
                    $Computer = "XLWUW-491S3W"
                    Enter-PSSession
                 }
            If ($Pick -eq 4)
                {
                    $Computer = "XLWUW-491S8S"
                    Enter-PSSession
                 }
            If ($Pick -eq 5)
                {
                    $Computer = "XLWUW-491S73"
                    Enter-PSSession
                 }
            If ($Pick -eq 6)
                {
                    $Computer = "XLWUW-491S96"
                    Enter-PSSession
                 }
            If ($Pick -eq 7)
                {
                    $Computer = "XLWUW-491S64"
                    Enter-PSSession
                 }
            If ($Pick -eq 8)
                {
                    $Computer = "XLWUL-410GP5"
                    Enter-PSSession
                 }
            If ($Pick -eq 9)
                {
                    $Computer = "XLWUW-491"
                    Enter-PSSession
                 }
             If ($Pick -eq 10)
                {
                    $Computer = "XLWUW-491S93"
                    Enter-PSSession
                 }
             If ($Pick -eq 11)
                {
                    $Computer = "XLWUW-491S38"
                    Enter-PSSession
                 }
              If ($Pick -eq 12)
                {
                    $Computer = "XLWUW-491S55"
                    Enter-PSSession
                 }
              If ($Pick -eq 13)
                {
                    $Computer = "XLWUW-491S8K"
                    Enter-PSSession
                 }
              If ($Pick -eq 14)
                {
                    $Computer = "XLWUW-491S5R"
                    Enter-PSSession
                 }
              If ($Pick -eq 15)
                {
                    $Computer = "XLWUW-491S5G"
                    Enter-PSSession
                 }
              If ($Pick -eq 16)
                {
                    $Computer = "XLWUW-491S8M"
                    Enter-PSSession
                 }
              If ($Pick -eq 17)
                {
                    $Computer = "XLWUW-491S5Y"
                    Enter-PSSession
                 }
              If ($Pick -eq 18)
                {
                    $Computer = "XLWUW-491S90"
                    Enter-PSSession
                 }
              If ($Pick -eq 19)
                {
                    $Computer = "XLWUW-491S7T"
                    Enter-PSSession
                 }
              If ($Pick -eq 20)
                {
                    $Computer = "XLWUW-491S5B"
                    Enter-PSSession
                 }
              If ($Pick -eq 21)
                {
                    $Computer = "XLWUW-491S40"
                    Enter-PSSession
                 }
              If ($Pick -eq 22)
                {
                    $Computer = "XLWUL-511KQF"
                    Enter-PSSession
                 }
              If ($Pick -eq 23)
                {
                    $Computer = "XLWUW-491S33"
                    Enter-PSSession
                 }
              If ($Pick -eq 24)
                {
                    $Computer = "XLWUW-AOCSD1"
                    Enter-PSSession
                 }
              If ($Pick -eq 25)
                {
                    $Computer = "XLWUW-491S4C"
                    Enter-PSSession
                 }  
              If ($Pick -eq 26)
                {
                    $Computer = "XLWUL-511KNP"
                    Enter-PSSession
                 } 
              If ($Pick -eq 27)
                {
                    $Computer = "XLWUW-491S3K"
                    Enter-PSSession
                 } 
              If ($Pick -eq 28)
                {
                    $Computer = "XLWUW-491S3H"
                    Enter-PSSession
                 } 
              If ($Pick -eq 29)
                {
                    $Computer = "XLWUW-6491S3B"
                    Enter-PSSession
                 } 
              If ($Pick -eq 30)
                {
                    $Computer = ""
                    Enter-PSSession
                 } 
              If ($Pick -eq 31)
                {
                    $Computer = ""
                    Enter-PSSession
                 } 
              If ($Pick -eq 32)
                {
                    $Computer = ""
                    Enter-PSSession
                 } 
              If ($Pick -eq 33)
                {
                    $Computer = "xlwuw-491s33"
                    Enter-PSSession
                 } 
              If ($Pick -eq 34)
                {
                    $Computer = "xlwuw-51m0nd5"
                    Enter-PSSession
                 }
              If ($Pick -eq 0)
                {
                  Write-Host          
                  $Computer = Read-Host "Computer"          
                  Enter-PSSession          
                 }
    }
    If ($Ans -eq 9)
    {
        ComputersPerBuilding
    }
    If ($Ans -eq 10)
    {
        ComputerAndUser
    } 
    If ($Ans -eq 11)
    {
        Write-Host
        $SorM = Read-Host "Single (s) or Multi (m) PCs?"
            If ($SorM -eq "s")
    {
    Write-Host " "
    Write-Host "0 - Enter Computer Name"
    Write-Host "1 - Alex"
    Write-Host "2 - Arnold"
    Write-Host "3 - Ballentine"
    Write-Host "4 - Barnett"
    Write-Host "5 - Ben"
    Write-Host "6 - Brown"
    Write-Host "7 - Cain"
    Write-Host "8 - Dossa"
    Write-Host "9 - Ed"
    Write-Host "10 - Gail"
    Write-Host "11 - Ginger" 
    Write-Host "12 - Goldman" 
    Write-Host "13 - Grainger"
    Write-Host "14 - Cunningham"
    Write-Host "15 - Hiserodt"
    Write-Host "16 - Johns"
    Write-Host "17 - Lafond"
    Write-Host "18 - Linde"
    Write-Host "19 - Lozada"
    Write-Host "20 - Mac"
    Write-Host "21 - Mowry"
    Write-Host "22 - Oster"
    Write-Host "23 - Pelletier"
    Write-Host "24 - Pelletier Admin"
    Write-Host "25 - Scheffrin"
    Write-Host "26 - SMSgt Larry"
    Write-Host "27 - Thrift"
    Write-Host "28 - Walden"
    Write-Host "29 - Worley"
        $Pick = Read-Host "Choose a Victim/Customer"   
            If ($Pick -eq 1)
                {
                    $Computer = "XLWUW-491S6B"
                    CDROM
                 }
            If ($Pick -eq 2)
                {
                    $Computer = "XLWUL-4422Z4"
                    CDROM
                 }
            If ($Pick -eq 3)
                {
                    $Computer = "XLWUW-491S3W"
                    CDROM
                 }
            If ($Pick -eq 4)
                {
                    $Computer = "XLWUW-491S8S"
                    CDROM
                 }
            If ($Pick -eq 5)
                {
                    $Computer = "XLWUW-491S73"
                    CDROM
                 }
            If ($Pick -eq 6)
                {
                    $Computer = "XLWUW-491S96"
                    CDROM
                 }
            If ($Pick -eq 7)
                {
                    $Computer = "XLWUW-491S64"
                    CDROM
                 }
            If ($Pick -eq 8)
                {
                    $Computer = "XLWUL-410GP5"
                    CDROM
                 }
            If ($Pick -eq 9)
                {
                    $Computer = "XLWUW-491"
                    CDROM
                 }
             If ($Pick -eq 10)
                {
                    $Computer = "XLWUW-491S93"
                    CDROM
                 }
             If ($Pick -eq 11)
                {
                    $Computer = "XLWUW-491S38"
                    CDROM
                 }
              If ($Pick -eq 12)
                {
                    $Computer = "XLWUW-491S55"
                    CDROM
                 }
              If ($Pick -eq 13)
                {
                    $Computer = "XLWUW-491S8K"
                    CDROM
                 }
              If ($Pick -eq 14)
                {
                    $Computer = "XLWUW-491S5R"
                    CDROM
                 }
              If ($Pick -eq 15)
                {
                    $Computer = "XLWUW-491S5G"
                    CDROM
                 }
              If ($Pick -eq 16)
                {
                    $Computer = "XLWUW-491S8M"
                    CDROM
                 }
              If ($Pick -eq 17)
                {
                    $Computer = "XLWUW-491S5Y"
                    CDROM
                 }
              If ($Pick -eq 18)
                {
                    $Computer = "XLWUW-491S90"
                    CDROM
                 }
              If ($Pick -eq 19)
                {
                    $Computer = "XLWUW-491S7T"
                    CDROM
                 }
              If ($Pick -eq 20)
                {
                    $Computer = "XLWUW-491S5B"
                    CDROM
                 }
              If ($Pick -eq 21)
                {
                    $Computer = "XLWUW-491S40"
                    CDROM
                 }
              If ($Pick -eq 22)
                {
                    $Computer = "XLWUL-511KQF"
                    CDROM
                 }
              If ($Pick -eq 23)
                {
                    $Computer = "XLWUW-491S33"
                    CDROM
                 }
              If ($Pick -eq 24)
                {
                    $Computer = "XLWUW-AOCSD1"
                    CDROM
                 }
              If ($Pick -eq 25)
                {
                    $Computer = "XLWUW-491S4C"
                    CDROM
                 }  
              If ($Pick -eq 26)
                {
                    $Computer = "XLWUL-511KNP"
                    CDROM
                 } 
              If ($Pick -eq 27)
                {
                    $Computer = "XLWUW-491S3K"
                    CDROM
                 } 
              If ($Pick -eq 28)
                {
                    $Computer = "XLWUW-491S3H"
                    CDROM
                 } 
              If ($Pick -eq 29)
                {
                    $Computer = "XLWUW-6491S3B"
                    CDROM
                 } 
              If ($Pick -eq 30)
                {
                    $Computer = ""
                    CDROM
                 } 
              If ($Pick -eq 31)
                {
                    $Computer = ""
                    CDROM
                 } 
              If ($Pick -eq 32)
                {
                    $Computer = ""
                    CDROM
                 } 
              If ($Pick -eq 33)
                {
                    $Computer = "xlwuw-491s33"
                    CDROM
                 } 
              If ($Pick -eq 34)
                {
                    $Computer = "xlwuw-51m0nd5"
                    CDROM
                 }
              If ($Pick -eq 0)
                {
                  Write-Host          
                  $Computer = Read-Host "Computer"          
                  CDROM          
                 }
    }

            If ($SorM -eq "m")
            {CDROM2}  
                }
    If ($Ans -eq 12)
    {
        RoboCopy
    } 
    If ($Ans -eq 13)
    {
    Write-Host " "
    Write-Host "0 - Enter Computer Name"
    Write-Host "1 - Alex"
    Write-Host "2 - Arnold"
    Write-Host "3 - Ballentine"
    Write-Host "4 - Barnett"
    Write-Host "5 - Ben"
    Write-Host "6 - Brown"
    Write-Host "7 - Cain"
    Write-Host "8 - Dossa"
    Write-Host "9 - Ed"
    Write-Host "10 - Gail"
    Write-Host "11 - Ginger" 
    Write-Host "12 - Goldman" 
    Write-Host "13 - Grainger"
    Write-Host "14 - Cunningham"
    Write-Host "15 - Hiserodt"
    Write-Host "16 - Johns"
    Write-Host "17 - Lafond"
    Write-Host "18 - Linde"
    Write-Host "19 - Lozada"
    Write-Host "20 - Mac"
    Write-Host "21 - Mowry"
    Write-Host "22 - Oster"
    Write-Host "23 - Pelletier"
    Write-Host "24 - Pelletier Admin"
    Write-Host "25 - Scheffrin"
    Write-Host "26 - SMSgt Larry"
    Write-Host "27 - Thrift"
    Write-Host "28 - Walden"
    Write-Host "29 - Worley"
        $Pick = Read-Host "Choose a Victim/Customer"   
            If ($Pick -eq 1)
                {
                    $Computer = "XLWUW-491S6B"
                    nslookup
                 }
            If ($Pick -eq 2)
                {
                    $Computer = "XLWUL-4422Z4"
                    nslookup
                 }
            If ($Pick -eq 3)
                {
                    $Computer = "XLWUW-491S3W"
                    nslookup
                 }
            If ($Pick -eq 4)
                {
                    $Computer = "XLWUW-491S8S"
                    nslookup
                 }
            If ($Pick -eq 5)
                {
                    $Computer = "XLWUW-491S73"
                    nslookup
                 }
            If ($Pick -eq 6)
                {
                    $Computer = "XLWUW-491S96"
                    nslookup
                 }
            If ($Pick -eq 7)
                {
                    $Computer = "XLWUW-491S64"
                    nslookup
                 }
            If ($Pick -eq 8)
                {
                    $Computer = "XLWUL-410GP5"
                    nslookup
                 }
            If ($Pick -eq 9)
                {
                    $Computer = "XLWUW-491"
                    nslookup
                 }
             If ($Pick -eq 10)
                {
                    $Computer = "XLWUW-491S93"
                    nslookup
                 }
             If ($Pick -eq 11)
                {
                    $Computer = "XLWUW-491S38"
                    nslookup
                 }
              If ($Pick -eq 12)
                {
                    $Computer = "XLWUW-491S55"
                    nslookup
                 }
              If ($Pick -eq 13)
                {
                    $Computer = "XLWUW-491S8K"
                    nslookup
                 }
              If ($Pick -eq 14)
                {
                    $Computer = "XLWUW-491S5R"
                    nslookup
                 }
              If ($Pick -eq 15)
                {
                    $Computer = "XLWUW-491S5G"
                    nslookup
                 }
              If ($Pick -eq 16)
                {
                    $Computer = "XLWUW-491S8M"
                    nslookup
                 }
              If ($Pick -eq 17)
                {
                    $Computer = "XLWUW-491S5Y"
                    nslookup
                 }
              If ($Pick -eq 18)
                {
                    $Computer = "XLWUW-491S90"
                    nslookup
                 }
              If ($Pick -eq 19)
                {
                    $Computer = "XLWUW-491S7T"
                    nslookup
                 }
              If ($Pick -eq 20)
                {
                    $Computer = "XLWUW-491S5B"
                    nslookup
                 }
              If ($Pick -eq 21)
                {
                    $Computer = "XLWUW-491S40"
                    nslookup
                 }
              If ($Pick -eq 22)
                {
                    $Computer = "XLWUL-511KQF"
                    nslookup
                 }
              If ($Pick -eq 23)
                {
                    $Computer = "XLWUW-491S33"
                    nslookup
                 }
              If ($Pick -eq 24)
                {
                    $Computer = "XLWUW-AOCSD1"
                    nslookup
                 }
              If ($Pick -eq 25)
                {
                    $Computer = "XLWUW-491S4C"
                    nslookup
                 }  
              If ($Pick -eq 26)
                {
                    $Computer = "XLWUL-511KNP"
                    nslookup
                 } 
              If ($Pick -eq 27)
                {
                    $Computer = "XLWUW-491S3K"
                    nslookup
                 } 
              If ($Pick -eq 28)
                {
                    $Computer = "XLWUW-491S3H"
                    nslookup
                 } 
              If ($Pick -eq 29)
                {
                    $Computer = "XLWUW-6491S3B"
                    nslookup
                 } 
              If ($Pick -eq 30)
                {
                    $Computer = ""
                    nslookup
                 } 
              If ($Pick -eq 31)
                {
                    $Computer = ""
                    nslookup
                 } 
              If ($Pick -eq 32)
                {
                    $Computer = ""
                    nslookup
                 } 
              If ($Pick -eq 33)
                {
                    $Computer = "xlwuw-491s33"
                    nslookup
                 } 
              If ($Pick -eq 34)
                {
                    $Computer = "xlwuw-51m0nd5"
                    nslookup
                 }
              If ($Pick -eq 0)
                {
                  Write-Host          
                  $Computer = Read-Host "Computer"          
                  nslookup          
                 }
    }
    If ($Ans -eq 14)
    {
    Write-Host " "
    Write-Host "0 - Enter Computer Name"
    Write-Host "1 - Alex"
    Write-Host "2 - Arnold"
    Write-Host "3 - Ballentine"
    Write-Host "4 - Barnett"
    Write-Host "5 - Ben"
    Write-Host "6 - Brown"
    Write-Host "7 - Cain"
    Write-Host "8 - Dossa"
    Write-Host "9 - Ed"
    Write-Host "10 - Gail"
    Write-Host "11 - Ginger" 
    Write-Host "12 - Goldman" 
    Write-Host "13 - Grainger"
    Write-Host "14 - Cunningham"
    Write-Host "15 - Hiserodt"
    Write-Host "16 - Johns"
    Write-Host "17 - Lafond"
    Write-Host "18 - Linde"
    Write-Host "19 - Lozada"
    Write-Host "20 - Mac"
    Write-Host "21 - Mowry"
    Write-Host "22 - Oster"
    Write-Host "23 - Pelletier"
    Write-Host "24 - Pelletier Admin"
    Write-Host "25 - Scheffrin"
    Write-Host "26 - SMSgt Larry"
    Write-Host "27 - Thrift"
    Write-Host "28 - Walden"
    Write-Host "29 - Worley"
        $Pick = Read-Host "Choose a Victim/Customer"   
            If ($Pick -eq 1)
                {
                    $Computer = "XLWUW-491S6B"
                    C$
                 }
            If ($Pick -eq 2)
                {
                    $Computer = "XLWUL-4422Z4"
                    C$
                 }
            If ($Pick -eq 3)
                {
                    $Computer = "XLWUW-491S3W"
                    C$
                 }
            If ($Pick -eq 4)
                {
                    $Computer = "XLWUW-491S8S"
                    C$
                 }
            If ($Pick -eq 5)
                {
                    $Computer = "XLWUW-491S73"
                    C$
                 }
            If ($Pick -eq 6)
                {
                    $Computer = "XLWUW-491S96"
                    C$
                 }
            If ($Pick -eq 7)
                {
                    $Computer = "XLWUW-491S64"
                    C$
                 }
            If ($Pick -eq 8)
                {
                    $Computer = "XLWUL-410GP5"
                    C$
                 }
            If ($Pick -eq 9)
                {
                    $Computer = "XLWUW-491"
                    C$
                 }
             If ($Pick -eq 10)
                {
                    $Computer = "XLWUW-491S93"
                    C$
                 }
             If ($Pick -eq 11)
                {
                    $Computer = "XLWUW-491S38"
                    C$
                 }
              If ($Pick -eq 12)
                {
                    $Computer = "XLWUW-491S55"
                    C$
                 }
              If ($Pick -eq 13)
                {
                    $Computer = "XLWUW-491S8K"
                    C$
                 }
              If ($Pick -eq 14)
                {
                    $Computer = "XLWUW-491S5R"
                    C$
                 }
              If ($Pick -eq 15)
                {
                    $Computer = "XLWUW-491S5G"
                    C$
                 }
              If ($Pick -eq 16)
                {
                    $Computer = "XLWUW-491S8M"
                    C$
                 }
              If ($Pick -eq 17)
                {
                    $Computer = "XLWUW-491S5Y"
                    C$
                 }
              If ($Pick -eq 18)
                {
                    $Computer = "XLWUW-491S90"
                    C$
                 }
              If ($Pick -eq 19)
                {
                    $Computer = "XLWUW-491S7T"
                    C$
                 }
              If ($Pick -eq 20)
                {
                    $Computer = "XLWUW-491S5B"
                    C$
                 }
              If ($Pick -eq 21)
                {
                    $Computer = "XLWUW-491S40"
                    C$
                 }
              If ($Pick -eq 22)
                {
                    $Computer = "XLWUL-511KQF"
                    C$
                 }
              If ($Pick -eq 23)
                {
                    $Computer = "XLWUW-491S33"
                    C$
                 }
              If ($Pick -eq 24)
                {
                    $Computer = "XLWUW-AOCSD1"
                    C$
                 }
              If ($Pick -eq 25)
                {
                    $Computer = "XLWUW-491S4C"
                    C$
                 }  
              If ($Pick -eq 26)
                {
                    $Computer = "XLWUL-511KNP"
                    C$
                 } 
              If ($Pick -eq 27)
                {
                    $Computer = "XLWUW-491S3K"
                    C$
                 } 
              If ($Pick -eq 28)
                {
                    $Computer = "XLWUW-491S3H"
                    C$
                 } 
              If ($Pick -eq 29)
                {
                    $Computer = "XLWUW-6491S3B"
                    C$
                 } 
              If ($Pick -eq 30)
                {
                    $Computer = ""
                    C$
                 } 
              If ($Pick -eq 31)
                {
                    $Computer = ""
                    C$
                 } 
              If ($Pick -eq 32)
                {
                    $Computer = ""
                    C$
                 } 
              If ($Pick -eq 33)
                {
                    $Computer = "xlwuw-491s33"
                    C$
                 } 
              If ($Pick -eq 34)
                {
                    $Computer = "xlwuw-51m0nd5"
                    C$
                 }
              If ($Pick -eq 0)
                {
                  Write-Host          
                  $Computer = Read-Host "Computer"          
                  C$          
                 }
    }
    If ($Ans -eq 15)
    {
    Write-Host " "
    Write-Host "0 - Enter Computer Name"
    Write-Host "1 - Alex"
    Write-Host "2 - Arnold"	
    Write-Host "3 - Ballentine"	
    Write-Host "4 - Barnett"	
    Write-Host "5 - Ben"	
    Write-Host "6 - Brown"	
    Write-Host "7 - Cain"	
    Write-Host "8 - Dossa"	
    Write-Host "9 - Ed"	
    Write-Host "10 - Gail"	
    Write-Host "11 - Ginger" 	
    Write-Host "12 - Goldman" 	
	Write-Host "13 - Grainger"
	Write-Host "14 - Cunningham"
	Write-Host "15 - Hiserodt"
	Write-Host "16 - Johns"
	Write-Host "17 - Lafond"
	Write-Host "18 - Linde"
	Write-Host "19 - Lozada"
	Write-Host "20 - Mac"
	Write-Host "21 - Mowry"
	Write-Host "22 - Oster"
	Write-Host "23 - Pelletier"
	Write-Host "24 - Pelletier Admin"
	Write-Host "25 - Scheffrin"
	Write-Host "26 - SMSgt Larry"
	Write-Host "27 - Thrift"
	Write-Host "28 - Walden"
	Write-Host "29 - Worley"
        $Pick = Read-Host "Choose a Victim/Customer"   	
            If ($Pick -eq 1)	
                {	
                    $Computer = "XLWUW-491S6B"
                    GPUpdate	
                 }	
            If ($Pick -eq 2)	
                {	
                    $Computer = "XLWUL-4422Z4"
                    GPUpdate	
                 }	
            If ($Pick -eq 3)	
                {	
                    $Computer = "XLWUW-491S3W"	
                    GPUpdate
                 }	
            If ($Pick -eq 4)	
                {	
                    $Computer = "XLWUW-491S8S"	
                    GPUpdate
                 }	
            If ($Pick -eq 5)	
                {	
                    $Computer = "XLWUW-491S73"	
                    GPUpdate
                 }	
            If ($Pick -eq 6)	
                {	
                    $Computer = "XLWUW-491S96"	
                    GPUpdate
                 }	
            If ($Pick -eq 7)	
                {	
                    $Computer = "XLWUW-491S64"	
                    GPUpdate
                 }	
            If ($Pick -eq 8)	
                {	
                    $Computer = "XLWUL-410GP5"	
                    GPUpdate
                 }	
            If ($Pick -eq 9)	
                {	
                    $Computer = "XLWUW-491"	
                    GPUpdate
                 }	
             If ($Pick -eq 10)	
                {	
                    $Computer = "XLWUW-491S93"	
                    GPUpdate
                 }	
             If ($Pick -eq 11)	
                {	
                    $Computer = "XLWUW-491S38"	
                    GPUpdate
                 }	
              If ($Pick -eq 12)	
                {	
                    $Computer = "XLWUW-491S55"	
                    GPUpdate
                 }	
              If ($Pick -eq 13)	
                {	
                    $Computer = "XLWUW-491S8K"	
                    GPUpdate
                 }
              If ($Pick -eq 14)	
                {	
                    $Computer = "XLWUW-491S5R"
                    GPUpdate	
                 }
              If ($Pick -eq 15)	
                {	
                    $Computer = "XLWUW-491S5G"	
                    GPUpdate
                 }
              If ($Pick -eq 16)	
                {	
                    $Computer = "XLWUW-491S8M"	
                    GPUpdate
                 }
              If ($Pick -eq 17)	
                {	
                    $Computer = "XLWUW-491S5Y"	
                    GPUpdate
                 }
              If ($Pick -eq 18)	
                {	
                    $Computer = "XLWUW-491S90"	
                    GPUpdate
                 }
              If ($Pick -eq 19)	
                {	
                    $Computer = "XLWUW-491S7T"	
                    GPUpdate
                 }
              If ($Pick -eq 20)	
                {	
                    $Computer = "XLWUW-491S5B"	
                    GPUpdate
                 }
              If ($Pick -eq 21)	
                {	
                    $Computer = "XLWUW-491S40"	
                    GPUpdate
                 }
              If ($Pick -eq 22)	
                {	
                    $Computer = "XLWUL-511KQF"	
                    GPUpdate
                 }
              If ($Pick -eq 23)	
                {	
                    $Computer = "XLWUW-491S33"	
                    GPUpdate
                 }
              If ($Pick -eq 24)	
                {	
                    $Computer = "XLWUW-AOCSD1"	
                    GPUpdate
                 }
              If ($Pick -eq 25)	
                {	
                    $Computer = "XLWUW-491S4C"	
                    GPUpdate
                 }  
              If ($Pick -eq 26)	
                {	
                    $Computer = "XLWUL-511KNP"	
                    GPUpdate
                 } 
              If ($Pick -eq 27)	
                {	
                    $Computer = "XLWUW-491S3K"	
                    GPUpdate
                 } 
              If ($Pick -eq 28)	
                {	
                    $Computer = "XLWUW-491S3H"	
                    GPUpdate
                 } 
              If ($Pick -eq 29)	
                {	
                    $Computer = "XLWUW-6491S3B"	
                    GPUpdate
                 } 
              If ($Pick -eq 30)	
                {	
                    $Computer = ""	
                    GPUpdate
                 } 
              If ($Pick -eq 31)	
                {	
                    $Computer = ""	
                    GPUpdate
                 } 
              If ($Pick -eq 32)	
                {	
                    $Computer = ""	
                    GPUpdate
                 } 
              If ($Pick -eq 33)	
                {	
                    $Computer = "xlwuw-491s33"
                    GPUpdate	
                 } 
              If ($Pick -eq 34)	
                {	
                    $Computer = "xlwuw-51m0nd5"	
                    GPUpdate
                 }
              If ($Pick -eq 0)	
                {	
                  Write-Host          
                  $Computer = Read-Host "Computer"          
                  GPUpdate          
                 }
    }
    If ($Ans -eq 16)
    {
    Write-Host " "
    Write-Host "0 - Enter Computer Name"
    Write-Host "1 - Alex"
    Write-Host "2 - Arnold"	
    Write-Host "3 - Ballentine"	
    Write-Host "4 - Barnett"	
    Write-Host "5 - Ben"	
    Write-Host "6 - Brown"	
    Write-Host "7 - Cain"	
    Write-Host "8 - Dossa"	
    Write-Host "9 - Ed"	
    Write-Host "10 - Gail"	
    Write-Host "11 - Ginger" 	
    Write-Host "12 - Goldman" 	
	Write-Host "13 - Grainger"
	Write-Host "14 - Cunningham"
	Write-Host "15 - Hiserodt"
	Write-Host "16 - Johns"
	Write-Host "17 - Lafond"
	Write-Host "18 - Linde"
	Write-Host "19 - Lozada"
	Write-Host "20 - Mac"
	Write-Host "21 - Mowry"
	Write-Host "22 - Oster"
	Write-Host "23 - Pelletier"
	Write-Host "24 - Pelletier Admin"
	Write-Host "25 - Scheffrin"
	Write-Host "26 - SMSgt Larry"
	Write-Host "27 - Thrift"
	Write-Host "28 - Walden"
	Write-Host "29 - Worley"
        $Pick = Read-Host "Choose a Victim/Customer"   	
            If ($Pick -eq 1)	
                {	
                    $Computer = "XLWUW-491S6B"	
                 }	
            If ($Pick -eq 2)	
                {	
                    $Computer = "XLWUL-4422Z4"	
                 }	
            If ($Pick -eq 3)	
                {	
                    $Computer = "XLWUW-491S3W"	
                 }	
            If ($Pick -eq 4)	
                {	
                    $Computer = "XLWUW-491S8S"	
                 }	
            If ($Pick -eq 5)	
                {	
                    $Computer = "XLWUW-491S73"	
                 }	
            If ($Pick -eq 6)	
                {	
                    $Computer = "XLWUW-491S96"	
                 }	
            If ($Pick -eq 7)	
                {	
                    $Computer = "XLWUW-491S64"	
                 }	
            If ($Pick -eq 8)	
                {	
                    $Computer = "XLWUL-410GP5"	
                 }	
            If ($Pick -eq 9)	
                {	
                    $Computer = "XLWUW-491"	
                 }	
             If ($Pick -eq 10)	
                {	
                    $Computer = "XLWUW-491S93"	
                 }	
             If ($Pick -eq 11)	
                {	
                    $Computer = "XLWUW-491S38"	
                 }	
              If ($Pick -eq 12)	
                {	
                    $Computer = "XLWUW-491S55"	
                 }	
              If ($Pick -eq 13)	
                {	
                    $Computer = "XLWUW-491S8K"	
                 }
              If ($Pick -eq 14)	
                {	
                    $Computer = "XLWUW-491S5R"	
                 }
              If ($Pick -eq 15)	
                {	
                    $Computer = "XLWUW-491S5G"	
                 }
              If ($Pick -eq 16)	
                {	
                    $Computer = "XLWUW-491S8M"	
                 }
              If ($Pick -eq 17)	
                {	
                    $Computer = "XLWUW-491S5Y"	
                 }
              If ($Pick -eq 18)	
                {	
                    $Computer = "XLWUW-491S90"	
                 }
              If ($Pick -eq 19)	
                {	
                    $Computer = "XLWUW-491S7T"	
                 }
              If ($Pick -eq 20)	
                {	
                    $Computer = "XLWUW-491S5B"	
                 }
              If ($Pick -eq 21)	
                {	
                    $Computer = "XLWUW-491S40"	
                 }
              If ($Pick -eq 22)	
                {	
                    $Computer = "XLWUL-511KQF"	
                 }
              If ($Pick -eq 23)	
                {	
                    $Computer = "XLWUW-491S33"	
                 }
              If ($Pick -eq 24)	
                {	
                    $Computer = "XLWUW-AOCSD1"	
                 }
              If ($Pick -eq 25)	
                {	
                    $Computer = "XLWUW-491S4C"	
                 }  
              If ($Pick -eq 26)	
                {	
                    $Computer = "XLWUL-511KNP"	
                 } 
              If ($Pick -eq 27)	
                {	
                    $Computer = "XLWUW-491S3K"	
                 } 
              If ($Pick -eq 28)	
                {	
                    $Computer = "XLWUW-491S3H"	
                 } 
              If ($Pick -eq 29)	
                {	
                    $Computer = "XLWUW-6491S3B"	
                 } 
              If ($Pick -eq 30)	
                {	
                    $Computer = ""	
                 } 
              If ($Pick -eq 31)	
                {	
                    $Computer = ""	
                 } 
              If ($Pick -eq 32)	
                {	
                    $Computer = ""	
                 } 
              If ($Pick -eq 33)	
                {	
                    $Computer = "xlwuw-491s33"	
                 } 
              If ($Pick -eq 34)	
                {	
                    $Computer = "xlwuw-51m0nd5"	
                 }        	
              If ($Pick -eq 0)	
                {	
                  Write-Host          
                  $Computer = Read-Host "Computer"          
                  StopProcess          
                 }
        Write-Host "1 - Outlook"
        Write-Host "2 - Word"
        Write-Host "3 - Excel"
        Write-Host "4 - Skype"
        Write-Host "5 - IE"
        Write-Host "6 - Chrome"
        Write-Host "7 - FireFox"
        Write-Host "8 - PowerShell"
        Write-Host "9 - DameWare"
        Write-Host "10 - Virtual"
        Write-Host "11 - AtHoc" 
        Write-Host "12 - Explorer"
        Write-Host "13 - PowerShell ISE" 
            $Which = Read-Host "Choose a beep-boop-beep to unbeep/boop"   
            If ($Which -eq 1)
                {
                    $service = "Outlook"
                    StopProcess
                 }
            If ($Which -eq 2)
                {
                    $service = "winword"
                    StopProcess
                 }
            If ($Which -eq 3)
                {
                    $service = "Excel"
                    StopProcess
                 }
            If ($Which -eq 4)
                {
                    $service = "lync"
                    StopProcess
                 }
            If ($Which -eq 5)
                {
                    $service = "iexplore"
                    StopProcess
                 }
            If ($Which -eq 6)
                {
                    $service = "Chrome"
                    StopProcess
                 }
            If ($Which -eq 7)
                {
                    $service = "FireFox"
                    StopProcess
                 }
            If ($Which -eq 8)
                {
                    $service = "Powershell"
                    StopProcess
                 }
            If ($Which -eq 9)
                {
                    $service = "DWRCC"
                    StopProcess
                 }
             If ($Which -eq 10)
                {
                    $service = "vmconnect"
                    StopProcess
                 }
             If ($Which -eq 11)
                {
                    $service = "AtHocUsaf"
                    StopProcess
                 }
              If ($Which -eq 12)
                {
                    $service = "explorer"
                    StopProcess
                 }
              If ($Which -eq 13)
                {
                    $service = "powershell_ise"
                    StopProcess
                 }          
                 }   
    }
    

Until ($Ans -eq 17)