<#
.NOTES
-------------------------------------
Name:    SplashScreen.ps1
Version: 1.0 - 03/25/2017
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------

.SYNOPSIS
This script launches a Winform Splash Screen on a Background Thread.
There are 2 functions, Show-SplashScreen & Close-SplashScreen.
--------------------------------------------------------------------------------------------

.DESCRIPTION
There are 2 functions, Show-SplashScreen & Close-SplashScreen, When calling 
Show-SplashScreen 4 parameters allow specifing the Calling Application Name,
An Image to Display, and Optionally the form background color & a switch which
causes the image to be used as a background image. By default the image is
displayed inside a PictureBox Control.
---------------------------------------------------------------------------------------- 
#>

<#
.NOTES
-------------------------------------------
Name:    Function Initialize-SplashRunspace
Version: 1.0 - 03/25/2017
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------------

.SYNOPSIS
Creates a Custom Runspace object used by Show-SplashScreen & Close-SplashScreen
----------------------------------------------------------------------------------------
.DESCRIPTION
Creates a Custom Runspace object used by Show-SplashScreen & Close-SplashScreen
A script level object $SplashRunspace is created when this module is loaded.
$SplashRunspace will normally not need to be reinitalized by the user, but may be
by issuing the following: $SplashRunspace = Initialize-SplashRunspace
---------------------------------------------------------------------------------------- 
.EXAMPLE
$SplashRunspace = Initialize-SplashRunspace
Initiales shared Runspace $SplashRunspace
#>
function Initialize-SplashRunspace{
$NewObj = New-Object -TypeName PSObject
#HashTable used for communication with the background thread.
$SSHash = [HashTable]::Synchronized(@{})
$SSHash.Flag = $True # Splash displayed while True
$SSHash.SplashIsLoaded = 0 # Value determines if Splash is displayed 0/1 = Off/On 
$PowerShell = [PowerShell]::Create()
$PowerShell.Runspace = [RunspaceFactory]::CreateRunspace()
$PowerShell.Runspace.Open()
#Pass HashTable used for communications
$PowerShell.Runspace.SessionStateProxy.SetVariable("SyncHash",$SSHash)

Add-Member -InputObject $NewObj -MemberType NoteProperty -Name HashTable  -Value $SSHash
Add-Member -InputObject $NewObj -MemberType NoteProperty -Name Powershell -Value $PowerShell
Add-Member -InputObject $NewObj -MemberType NoteProperty -Name Handle     -Value 0
return $NewObj
}
 
<#
.NOTES
-------------------------------------
Name:    Function Show-SplashScreen
Version: 1.0 - 03/25/2017
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------

.SYNOPSIS
This function Displays a Splash Screen on a Background Thread.
----------------------------------------------------------------------------------------
.DESCRIPTION
This function has 4 parameters allowing specifing the Calling Application Name,
An Image to Display, and Optionally the form background color, & a switch which causes 
the image to be used as a background image. By default the image is displayed inside a
PictureBox Control.
---------------------------------------------------------------------------------------- 
.Parameter AppName - Alias: Name
Optional, Name of the application to be displayed - Defaults to "Application".

.PARAMETER Image - Alias: Img
Required, A System.Drawing.Image to display on the splash screen.

.PARAMETER FormBackColor - Alias: FBC
Optional form background color

.PARAMETER BgImage - Alias: BGI
Optional, Switch - If present the image is displayed as a background image, 
otherwise the image is displayed in a PictureBox.

.EXAMPLE
Show-SplashScreen -AppName <Name> -Image <Image>
Displays the Splash Screen using a PictureBox
#>
function Show-SplashScreen{
param(
	[Parameter(Mandatory = $False)][Alias('Name')][String]$AppName = "Application",
	[Parameter(Mandatory = $True)][Alias('Img')][System.Drawing.Image]$Image,
	[Parameter(Mandatory = $False)][Alias('FBC')][System.Drawing.Color]$FormBackColor=[Drawing.SystemColors]::Control,
	[Parameter(Mandatory = $False)][Alias('BGI')][Switch]$BgImage)

if($SplashRunspace.HashTable.SplashIsLoaded -gt 0)
	{return} #Abort Call Already Active

$SplashScript = {
param(
	[Parameter(Mandatory = $False)][Alias('Name')][String]$AppName = "Application",
	[Parameter(Mandatory = $True)][Alias('Img')][System.Drawing.Image]$Image,
	[Parameter(Mandatory = $False)][Alias('FBC')][System.Drawing.Color]$FormBackColor=[Drawing.SystemColors]::Control,
	[Parameter(Mandatory = $False)][Alias('BGI')][Switch]$BgImage)

while($SyncHash.Flag){
	if($SyncHash.SplashIsLoaded -eq 0) # First Cycle of While
		{
		Add-Type -AssemblyName System.Windows.Forms

		$Splash = New-Object -TypeName System.Windows.Forms.Form
		$Timer = New-Object -TypeName System.Windows.Forms.Timer
		$ProgressBar = New-Object -TypeName System.Windows.Forms.ProgressBar

		$Splash.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 475, 325
		$Splash.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
		$Splash.Font = New-Object -TypeName System.Drawing.Font `
			-ArgumentList "Arial", 12,[System.Drawing.FontStyle]::Regular
		$Splash.ControlBox = $False
		$Splash.BackColor = $FormBackColor
		if($BgImage.IsPresent){
			$Splash.BackgroundImage = $Image
			$Splash.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Zoom
		}
		$Splash.FormBorderStyle = [System.Windows.Forms.BorderStyle]::None
		$Splash.ShowInTaskbar = $False
		$SharedWidth = $Splash.Width - 25

		$Timer.Interval = 100
		$Timer_Tick = {
			$Timer.Enabled = $False
			if($SyncHash.Flag -eq $False){$Splash.Close()}
			if($ProgressBar.Value -eq $ProgressBar.Maximum){$ProgressBar.Value = 0}
			$ProgressBar.Value++
			$ProgressBar.Refresh()
			[System.Windows.Forms.Application]::DoEvents()
			$Timer.Enabled = $True
		}
		$Timer.Add_Tick($Timer_Tick)
		$Timer.Enabled = $True
		$Timer.Start()

		$ProgressBar.Style = 'Continuous'
		$ProgressBar.Parent = $Splash
		$ProgressBar.Minimum = 0
		$ProgressBar.Maximum = 100
		$ProgressBar.Enabled = $True
		$ProgressBar.Visible = $True
		$ProgressBar.Size = New-Object -TypeName System.Drawing.Size `
			-ArgumentList $SharedWidth,30
		$ProgressBar.Location = New-Object -TypeName System.Drawing.Point `
			-ArgumentList 13,($Splash.Height-($ProgressBar.Height+10))
		$ProgressBar.BackColor = [Drawing.SystemColors]::Control
		$Splash.Controls.Add($ProgressBar)
    
		$LblMsg = New-Object -TypeName System.Windows.Forms.Label
		$LblMsg.Parent = $Splash
		$LblMsg.Size = New-Object -TypeName System.Drawing.Size `
			-ArgumentList $SharedWidth,20
		$LblMsg.Location = New-Object -TypeName System.Drawing.Point `
			-ArgumentList 13,($ProgressBar.top-($LblMsg.Height+10))
		$LblMsg.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
		$LblMsg.BackColor = [Drawing.SystemColors]::Control
		$LblMsg.Text = "{0} Loading, Please Wait ..." -f $AppName
		$LblMsg.Visible = $True
    
		if(!$BgImage.IsPresent){
			$PicBox = New-Object -TypeName System.Windows.Forms.PictureBox
			$PicBox.Parent = $Splash
			$PicBox.Size = New-Object -TypeName System.Drawing.Size `
				-ArgumentList $SharedWidth,($LblMsg.Top-25)
			$PicBox.Location = New-Object -TypeName System.Drawing.Point `
				-ArgumentList 13,13
			$PicBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
			$PicBox.Image = $image
			$PicBox.Visible = $True
			$PicBox.BackColor = [Drawing.SystemColors]::Control
			$PicBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
		}
   
		$InitialFormWindowState = $Splash.WindowState
		$Form_Load_StateCorrection = {$Splash.WindowState = $InitialFormWindowState}
		$Splash.Add_Load($Form_Load_StateCorrection)
		$Splash.Add_Shown({$Splash.Activate()})
		$SyncHash.SplashIsLoaded++
		$Splash.ShowDialog()
		}     
	}
}

#Pass Script & Parameters to runspace
[void]$SplashRunspace.Powershell.AddScript($SplashScript)
[void]$SplashRunspace.Powershell.AddParameter("AppName",$AppName)
[void]$SplashRunspace.Powershell.AddParameter("Image",$Image)
[void]$SplashRunspace.Powershell.AddParameter("FormBackColor",$FormBackColor)
if($BgImage.IsPresent)
	{[void]$SplashRunspace.Powershell.AddParameter("BgImage")}

$SplashRunspace.Handle = $SplashRunspace.PowerShell.BeginInvoke()
}

<#
.NOTES
-------------------------------------
Name:    Function Close-SplashScreen
Version: 1.0 - 03/25/2017
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------

.SYNOPSIS
This function Closes a Splash Screen on a Background Thread & Dispose of Thread.
----------------------------------------------------------------------------------------
.DESCRIPTION
This function Closes a Splash Screen 
---------------------------------------------------------------------------------------- 
.EXAMPLE
Close-SplashScreen
Displays the Splash Screen using a PictureBox
#>
function Close-SplashScreen{

$SplashRunspace.HashTable.Flag = $False	#Close Screen

if($SplashRunspace.Handle.IsCompleted)
	{#Dispose of Thread
	$SplashRunspace.PowerShell.EndInvoke($SplashRunspace.Handle)
	$SplashRunspace.PowerShell.Dispose()
	$SplashRunspace = $Null
	}    
}

<#
Initialize the Custom Runspace Object used by
Show-SplashScreen & Close-SplashScreen
reissue the following if you need to restart
the Splash Screen in a new Runspace
#>
$SplashRunspace = Initialize-SplashRunspace

$Demo = 0
if($Demo -gt 0){
	#region Add Custom DLL - Data Type
	<#BlueflameDynamics.IconTools Class#>
	if(!(Get-Module -Name Exists)){Import-Module -Name .\Exists.ps1}
	$DLLPath = ".\IconTools.dll"
	if((Exists -Mode File -Location $DLLPath) -eq $True){
		Add-Type -Path $DLLPath
		Start-Sleep -Seconds 1
		$SSImage = [BlueflameDynamics.IconTools]::ExtractIcon(
			$env:WinDir+"\system32\imageres.dll",311,64)
		}
	# Default Image (Powershell Icon) used for Demo/Testing 
	#endregion
	Show-SplashScreen -Image $SSImage -AppName "SplashScreen Demo" -FBC ([Drawing.Color]::CornflowerBlue)
	Start-sleep -Seconds 30
	Close-SplashScreen
	}