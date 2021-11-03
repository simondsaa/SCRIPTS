<#
.NOTES
Name:	WinFormsLibrary.ps1
Author:  Randy Turner
Version: 1.0
Date:    06/05/2014

.SYNOPSIS
Provides a wrapper for utility fumctions used with WinForms.
#>

#region Add BlueflameDynamics.IconTools Class
if(!(Get-Module -Name Exists)){Import-Module -Name .\Exists.ps1}
$DLLPath = ".\IconTools.dll"
if((Exists -Mode File -Location $DLLPath) -eq $True){Add-Type -Path $DLLPath}
#endregion

<#
.NOTES
Name:    Get-Anchor Function
Author:  Randy Turner
Version: 1.0
Date:    06/05/2014

.SYNOPSIS
Provides a wrapper for fumction used to get a WinForm Anchor value.

.PARAMETER Top
Alias: T
Optional, causes the TOP Anchor to be included.

.PARAMETER Bottom
Alias: B
Optional, causes the BOTTOM Anchor to be included.

.PARAMETER Left
Alias: L
Optional, causes the LEFT Anchor to be included.

.PARAMETER Right
Alias: R
Optional, causes the RIGHT Anchor to be included.

.EXAMPLE
$Textbox1.Anchor = Get-Anchor -T -L -B -R
This example returns an Anchor value for all four Anchors.

.EXAMPLE
$Textbox1.Anchor = Get-Anchor -T -L
This example returns an Anchor value for the Top & Left Anchors.

.EXAMPLE
$Button1.Anchor = Get-Anchor -B -R
This example returns an Anchor value for the Bottom & Right Anchors.
#>
function Get-Anchor{
	param(
		[Parameter(Mandatory=$False)][Alias('T')][Switch]$Top,
		[Parameter(Mandatory=$False)][Alias('B')][Switch]$Bottom,
		[Parameter(Mandatory=$False)][Alias('L')][Switch]$Left,
		[Parameter(Mandatory=$False)][Alias('R')][Switch]$Right)
    
	$Anchors = @(0,0,0,0)
	$Anchors[0] = if($Top    -eq $True){[Windows.Forms.AnchorStyles]::Top}
	$Anchors[1] = if($Bottom -eq $True){[Windows.Forms.AnchorStyles]::Bottom} 
	$Anchors[2] = if($Left   -eq $True){[Windows.Forms.AnchorStyles]::Left}
	$Anchors[3] = if($Right  -eq $True){[Windows.Forms.AnchorStyles]::Right}     
	return [Windows.Forms.AnchorStyles]`
		$($Anchors[0] -bor $Anchors[1] -bor $Anchors[2] -bor $Anchors[3])
}

<#
.NOTES
Name:    Get-Cursor Function
Author:  Randy Turner
Version: 1.0
Date:    06/05/2014

.SYNOPSIS
Provides a wrapper for fumction used to set a WinForm Cursor

.PARAMETER Mode
Required, Cursor type to return 

.EXAMPLE
$MainForm.Cursor = Get-Cursor -Mode AppStarting
This example returns an AppStarting Cursor

.EXAMPLE
$MainForm.Cursor = Get-Cursor -Mode WaitCursor
This example returns the WaitCursor

.EXAMPLE
$MainForm.Cursor = Get-Cursor Default
This example returns the Default Cursor
#>
function Get-Cursor{
	param(
		[Parameter(Mandatory=$True)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet("AppStarting","Arrow","Cross","Default","Hand","Help","HSplit",
					 "IBeam","No","NoMove2D","NoMoveHoriz","NoMoveVert","PanEast",
					 "PanNE","PanNorth","PanNW","PanSE","PanWest","PanSouth","PanSW",
					 "SizeAll","SizeNESW","SizeWE","UpArrow","VSplit","WaitCursor")]
		[String]$Mode)

	$Modes = @("AppStarting","Arrow","Cross","Default","Hand","Help","HSplit","IBeam","No","NoMove2D",
			   "NoMoveHoriz","NoMoveVert","PanEast","PanNE","PanNorth","PanNW","PanSE","PanWest",
			   "PanSouth","PanSW","SizeAll","SizeNESW","SizeWE","UpArrow","VSplit","WaitCursor") 

	switch([array]::IndexOf($Modes, $Mode)){ 
		#Set Cursor
		00 {[Windows.Forms.Cursors]::AppStarting}
		01 {[Windows.Forms.Cursors]::Arrow}
		02 {[Windows.Forms.Cursors]::Cross}
		04 {[Windows.Forms.Cursors]::Hand}
		05 {[Windows.Forms.Cursors]::Help}
		06 {[Windows.Forms.Cursors]::HSplit}
		07 {[Windows.Forms.Cursors]::IBeam}
		08 {[Windows.Forms.Cursors]::No}
		09 {[Windows.Forms.Cursors]::NoMove2D}
		10 {[Windows.Forms.Cursors]::NoMoveHoriz}
		11 {[Windows.Forms.Cursors]::NoMoveVert}
		12 {[Windows.Forms.Cursors]::PanEast}
		13 {[Windows.Forms.Cursors]::PanNE}
		14 {[Windows.Forms.Cursors]::PanNorth}
		15 {[Windows.Forms.Cursors]::PanNW}
		16 {[Windows.Forms.Cursors]::PanSE}
		17 {[Windows.Forms.Cursors]::PanSouth}
		18 {[Windows.Forms.Cursors]::PanSW}
		19 {[Windows.Forms.Cursors]::PanWest}
		20 {[Windows.Forms.Cursors]::SizeAll}
		21 {[Windows.Forms.Cursors]::SizeNESW}
		22 {[Windows.Forms.Cursors]::SizeWE}
		23 {[Windows.Forms.Cursors]::UpArrow}
		24 {[Windows.Forms.Cursors]::VSplit}
		25 {[Windows.Forms.Cursors]::WaitCursor}
		Default {[Windows.Forms.Cursors]::Default} #03
	}
}

<#
.NOTES
Name:    Invoke-FontDialog Function
Author:  Randy Turner
Version: 1.0
Date:    06/05/2014

.SYNOPSIS
Provides a wrapper for fumction used to Invoke a WinForm FontDialog

.PARAMETER Control
Required, Windows Control to interact with the FontDialog

.PARAMETER MinSize
Optional, Minimum font size in points, Defaults to 9

.PARAMETER MaxSize
Optional, Maximum font size in points, Defaults to 24

.PARAMETER Showcolor
Optional, Switch if present allows Font color

.PARAMETER FontMustExist
Optional, Switch if present the Font must exist on the host

.PARAMETER FixedPitchOnly
Optional, Switch if present only Fixed Pitch Fonts will be listed

.PARAMETER AllowSimulations
Optional, Switch if present allows graphics device interface (GDI) font simulations

.PARAMETER AllowVerticalFonts
Optional, Switch if present allows listing Vertical Fonts 

.PARAMETER AllowVectorFonts
Optional, Switch if present allows listing Vector Fonts

.PARAMETER ShowEffects
Optional, Switch if present allows the user to specify strikethrough, underline, and text color options

.EXAMPLE
$MainForm.Font = Invoke-FontDialog -Control $MainForm -FixedPitchOnly
This example returns a Selected Fixed Pitch font or leaves the font unchanged if the dialog is cancelled

.EXAMPLE
$ListView.Font = Invoke-FontDialog -Control $ListView -FixedPitchOnly -ShowEffects -MinSize 6
This example returns a Selected Fixed Pitch font, allow effects, & set the minimum font size to 6pts 
#>
function Invoke-FontDialog{
	param(
		[Parameter(Mandatory = $True)][Windows.Forms.Control]$Control,
		[Parameter(Mandatory = $False)][Int]$MinSize = 9,
		[Parameter(Mandatory = $False)][Int]$MaxSize = 24,
		[Parameter(Mandatory = $False)][Switch]$Showcolor,
		[Parameter(Mandatory = $False)][Switch]$FontMustExist,
		[Parameter(Mandatory = $False)][Switch]$FixedPitchOnly,
		[Parameter(Mandatory = $False)][Switch]$AllowSimulations,
		[Parameter(Mandatory = $False)][Switch]$AllowVerticalFonts,
		[Parameter(Mandatory = $False)][Switch]$AllowVectorFonts,
		[Parameter(Mandatory = $False)][Switch]$ShowEffects)

	$FontDialog = New-Object -TypeName Windows.Forms.FontDialog
	$FontDialog.Showcolor = $Showcolor
	$FontDialog.FontMustExist = $FontMustExist
	$FontDialog.FixedPitchOnly = $FixedPitchOnly
	$FontDialog.AllowSimulations = $AllowSimulations
	$FontDialog.AllowVerticalFonts = $AllowVerticalFonts
	$FontDialog.AllowVectorFonts = $AllowVectorFonts
	$FontDialog.ShowEffects = $ShowEffects
	$FontDialog.minSize = $MinSize
	$FontDialog.maxSize = $MaxSize
	$FontDialog.Font = $Control.Font
	$FontDialog.Font = $Control.Font
	$RV = $FontDialog.ShowDialog()
	if($RV -eq "OK"){
		$Control.Font = $FontDialog.Font
		$Control.Refresh
	}
	return $RV
}

<#
.NOTES
Name:    Show-AboutForm Function
Author:  Randy Turner
Version: 1.0
Date:    06/05/2014

.SYNOPSIS
Provides a wrapper for fumction used to Display an AboutBox

.PARAMETER AppName
Required, Application Name to be used within the AboutBox

.PARAMETER AboutText
Required, Text to be used within the AboutBox's RichTextBox

.PARAMETER FormWidth
Optional, Width of AboutBox Form, Default is 415

.PARAMETER FormHeight
Optional, Height of AboutBox Form, Default is 270

.PARAMETER DetectUrls
Optional, Switch if present causes the URLs within the About Text to be detected, marked, & Clickable

.EXAMPLE
$About_Click ={
$AboutText = ("<Some Help Test>") 
Show-AboutForm -AppName "AboutBox Test" -DetectUrls -AboutText $AboutText
}
This example displays the default About
#>
function Show-AboutForm{
	param (
		[Parameter(Mandatory = $True)][Alias('N')][String]$AppName,
		[Parameter(Mandatory = $True)][Alias('T')][String]$AboutText,
		[Parameter(Mandatory = $False)][Alias('W')][Int]$FormWidth = 422,
		[Parameter(Mandatory = $False)][Alias('H')][Int]$FormHeight = 270,
		[Parameter(Mandatory = $False)][Alias('URL')][Switch]$DetectUrls)

	Add-Type -A "System.Windows.Forms"
	#Add objects for About
	$FormAbout = New-Object -TypeName Windows.Forms.Form
	$RtbAbout = New-Object -TypeName Windows.Forms.RichTextBox
	$InitialFormWindowState = New-Object -TypeName Windows.Forms.FormWindowState
	
	#About Form
	$FormAbout.Name = "FormAbout"
	$FormAbout.AutoScroll = $True
	$FormAbout.ClientSize = New-Object -TypeName Drawing.Size($FormWidth, $FormHeight)
	$FormAbout.DataBindings.DefaultDataSourceUpdateMode = 0
	$FormAbout.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedSingle
	$FormAbout.StartPosition = [Windows.Forms.FormStartPosition]::CenterParent
	$FormAbout.Text = " About " + $AppName
	$FormAbout.Icon = [BlueflameDynamics.IconTools]::ExtractIcon("imageres.dll", 76, 24)
	$FormAbout.MaximizeBox = `
	$FormAbout.MinimizeBox = $False

	$RtbAbout.Name = "RtbAbout"
	$RtbAbout.Size = New-Object -TypeName Drawing.Size(($FormWidth-13), ($FormHeight-13))
	$RtbAbout.Location = New-Object -TypeName Drawing.Point(13, 13)
	$RtbAbout.Anchor = Get-Anchor -T -L -B -R
	$RtbAbout.BackColor = [Drawing.Color]::FromArgb(255, 240, 240, 240)
	$RtbAbout.BorderStyle = 0
	$RtbAbout.DataBindings.DefaultDataSourceUpdateMode = 0
	$RtbAbout.DetectUrls = $DetectUrls
	$RtbAbout.ReadOnly = $True
	$RtbAbout.Cursor = Get-Cursor -Mode Default
	$RtbAbout.TabIndex = 0
	$RtbAbout.TabStop = $False
	$RtbAbout.Text = -Join $FormAbout.Text, $AboutText
	
	#Handles clicking the links in about form
	$RtbAbout.add_LinkClicked({ Start-Process -FilePath $_.LinkText })
	$FormAbout.Controls.Add($RtbAbout)
	[Void]$FormAbout.ShowDialog()
}

<#
.NOTES
Name:    Show-InputBox Function
Author:  Randy Turner
Version: 1.0
Date:    06/05/2014

.SYNOPSIS
Provides a wrapper for fumction used to display an InputBox & Return the entered text
A $null is returned upon cancellation

.PARAMETER Prompt (Alias <P>)
Required, Text to Prompt for Input

.PARAMETER Title (Alias <T>)
Optional, Text of InputBox Title

.PARAMETER Default (Alias <DV>)
Optional, Text of InputBox Default Value

.EXAMPLE
$RV = Show-InputBox -Prompt "Find?" -Title "Search";if ($RV -ne "") { Find-ListViewItem $ListView[0] $RV }
This example prompts for input of a search value
#>
function Show-InputBox{
	param (
		[Parameter(Mandatory = $True)][Alias('P')][String]$Prompt,
		[Parameter(Mandatory = $False)][Alias('T')][String]$Title = "",
		[Parameter(Mandatory = $False)][Alias('DV')][String]$Default = "")
	
	If ($Title.Length -eq 0) { $Title = " " }
	
	return [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title, $Default)
}

<#
.NOTES
Name:    Show-MessageBox Function
Author:  Randy Turner
Version: 1.0
Date:    06/05/2014

.SYNOPSIS
Provides a wrapper for fumction used to Display a MessageBox & get the button selected

.PARAMETER Msg - Alias (M)
Required, Message Text

.PARAMETER Msg - Alias (T)
Optional, Form Title Text

.PARAMETER MessageBoxStyle - Alias (S)
Optional, Message Box Style with optional buttons

.PARAMETER MessageBoxIcon - Alias (I)
Optional, Message Box Icon name

.EXAMPLE
$RV = Show-MessageBox -Msg "This is a Test!" -Title "MyMsgBox" -MessageBoxStyle YesNoCancel -MessageBoxIcon Informational
This example displays a MessageBox returns the selected button
#>
function Show-MessageBox{
	param(
	[Parameter(Mandatory=$True)][Alias('M')][String]$Msg,
	[Parameter(Mandatory=$False)][Alias('T')][String]$Title = "",
	[Parameter(Mandatory=$False)][Alias('S')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet("OkOnly","OkCancel","AbortRetryIgnore","YesNoCancel","YesNo","RetryCancel")]
		[String]$MessageBoxStyle = "OkOnly",
	[Parameter(Mandatory=$False)][Alias('I')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet("NoIcon","Critical","Question","Warning","Informational")]
		[String]$MessageBoxIcon = "NoIcon")

	$MessageBoxStyles = @("OkOnly","OkCancel","AbortRetryIgnore","YesNoCancel","YesNo","RetryCancel")
	$MessageBoxIcons  = @("NoIcon","Critical","Question","Warning","Informational")

	#Set MessageBox Style
	$Type = [array]::IndexOf($MessageBoxStyles, $MessageBoxStyle)
	
	#Set MessageBox Icon
	$Icon = [array]::IndexOf($MessageBoxIcons, $MessageBoxIcon) * 16
	
	#Loads the WinForm Assembly
	Add-Type -A "System.Windows.Forms"

	#Display the message with input
	$Answer = [Windows.Forms.MessageBox]::Show($Msg,$Title,$Type,$Icon)
	
	#Return Answer
	Return $Answer
}

<#
.NOTES
Name:    Show-HelpForm Function
Author:  Randy Turner
Version: 1.0
Date:    06/05/2014

.SYNOPSIS
Provides a wrapper for fumction used to Display a Help Window

.PARAMETER AppName
Required, Application Name to be used within the AboutBox

.PARAMETER HelpText
Required, Text to be used within the HelpForm's RichTextBox

.PARAMETER FormWidth
Optional, Width of Help Form, Default is 510

.PARAMETER FormHeight
Optional, Height of Help Form, Default is 560

.PARAMETER DetectUrls - Alias: URL
Optional, Switch if present causes the URLs within the About Text to be detected, marked, & Clickable

.PARAMETER ReadHelpFile - Alias: File
Optional, Switch if present causes the HelpText parameter to be treated as a text filename
The file is imported and displayed in the Help Form.

.EXAMPLE
$Help_Click ={
$HelpText = ("<Some Help Test>") 
Show-HelpForm -AppName "Help Test" -DetectUrls -HelpText $HelpText
}
This example displays the default About
#>
function Show-HelpForm{

	param (
			[Parameter(Mandatory = $True)][String]$AppName,
			[Parameter(Mandatory = $True)][String]$HelpText,
			[Parameter(Mandatory = $False)][Int]$FormWidth = 510,
			[Parameter(Mandatory = $False)][Int]$FormHeight = 560,
			[Parameter(Mandatory = $False)][Alias('URL')][Switch]$DetectUrls,
			[Parameter(Mandatory = $False)][Alias('File')][Switch]$ReadHelpFile,
			[Parameter(Mandatory = $False)][Alias('Read')][Switch]$ReadText)

	Add-Type -A "System.Windows.Forms"

	#region Utility Functions
	function Set-MenuItem{
		param(
			[Parameter(Mandatory = $True)][Array]$Labels,
			[Parameter(Mandatory = $True)][Array]$MenuItems,
			[Parameter(Mandatory = $True)][Drawing.Size]$ItemSize,
			[Parameter(Mandatory = $True)][String]$ItemPrefix,
			[Parameter(Mandatory = $False)][Bool]$SetSizeOff = $False,
			[Parameter(Mandatory = $False)][Array]$HotKeys = @())

		for($C = 0; $C -le $MenuItems.GetUpperBound(0); $C++){
			$MenuItems[$C].Name = -join ($ItemPrefix, ($C + 1))
			$MenuItems[$C].Text = $Labels[$C]
			if($SetSizeOff -eq $False){
				$MenuItems[$C].Size = $ItemSize
				$MenuItems[$C].ShortcutKeys =`
				Get-ShortcutKey -Mode $(if($Hotkeys.Length -eq 0) {$Labels[$C]} else {$HotKeys[$C]})
			}
		}
	}
	#endregion

	#region Add objects for Help
	$FrmHelp = New-Object Windows.Forms.Form
	$rtbHelp = New-Object Windows.Forms.RichTextBox
	$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
	$Help_MainMenu = New-Object -TypeName Windows.Forms.MenuStrip
	$Help_MainMenuItems = @(for ($C = 1; $C -le 1; $C++) {New-Object -TypeName Windows.Forms.ToolStripMenuItem})
	$Help_ToolMenuItems = @(for ($C = 1; $C -le 1; $C++) {New-Object -TypeName Windows.Forms.ToolStripMenuItem})
	#endregion 

	#region Help form
	$frmHelp.AutoScroll = $True 
	$frmHelp.DataBindings.DefaultDataSourceUpdateMode = 0
	$System_Drawing_Size = New-Object Drawing.Size($FormWidth,$FormHeight)
	$frmHelp.MinimumSize = $System_Drawing_Size
	#$frmHelp.MaximumSize = $System_Drawing_Size
	$frmHelp.Name = "frmHelp"
	$frmHelp.StartPosition = 1
	$frmHelp.Text = -join $AppName," - Help"
	$frmHelp.FormBorderStyle = [Windows.Forms.FormBorderStyle]::Sizable
	$frmHelp.StartPosition = [Windows.Forms.FormStartPosition]::CenterParent
	$frmHelp.Icon = [BlueflameDynamics.IconTools]::ExtractIcon("imageres.dll", 94, 16)
	$frmHelp.MaximizeBox = $False
	$FrmHelp.Controls.Add($Help_MainMenu)
	#endregion

	#region Main Menu
	$Help_MainMenu.Name = "Help_MainMenu"
	$Help_MainMenu.Visible = $True
	$Help_MainMenu.Size = New-Object -TypeName Drawing.Size -ArgumentList 220,30
	$Help_MainMenu.Items.AddRange($Help_MainMenuItems)
	$Help_MainMenuItemsSize = New-Object -TypeName Drawing.Size -ArgumentList 219,22
	$Help_MainMenuText = @("Tools")
	Set-MenuItem $Help_MainMenuText $Help_MainMenuItems $Help_MainMenuItemsSize "Help_MainMenuItem" $True
	

	$Help_ToolMenuText = @("Font Settings")
	Set-MenuItem $Help_ToolMenuText $Help_ToolMenuItems $Help_MainMenuItemsSize "Help_ToolMenuItem" $False
	$Help_MainMenuItems[0].DropDownItems.AddRange($Help_ToolMenuItems)
	$Help_FontSettings_Click = {Invoke-FontDialog -Control $rtbHelp -FontMustExist -AllowSimulations -AllowVectorFonts}
	$Help_ToolMenuItems[0].Add_Click($Help_FontSettings_Click)
	#endregion

	#region Rich Textbox
	$rtbHelp.Size = New-Object Drawing.Size(($FormWidth-45),($FormHeight*.78))
	$rtbHelp.Anchor = Get-Anchor -T -L -B -R 
	$rtbHelp.BackColor = [Drawing.Color]::FromArgb(255,240,240,240)
	$rtbHelp.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
	$rtbHelp.DataBindings.DefaultDataSourceUpdateMode = 0
	$rtbHelp.Location = New-Object Drawing.Point(13,30)
	$rtbHelp.Name = "rtbHelp"
	$rtbHelp.ReadOnly = $True
	$rtbHelp.SelectionProtected = $True
	$rtbHelp.Cursor = Get-Cursor Default
	$rtbHelp.TabIndex = 0
	$rtbHelp.TabStop = $False
	$rtbHelp.DetectUrls = $DetectUrls

	if($ReadHelpFile.IsPresent){
		if((Exists -Mode File -Location $Helptext) -eq $False){
			Show-MessageBox `
				-Msg ("Help File: {0} Not Found!" -f $HelpText) `
				-Title 'Help File Not Found!' -MessageBoxStyle OkOnly -MessageBoxIcon Warning
		return $Null
		}
	$Lines = Get-Content -Path $Helptext
	$HelpText = ''
	foreach($Line in $Lines){$HelpText+="{0}`r`n" -f $Line}
	}

	$rtbHelp.Text = -join $AppName,' - Help',$HelpText

	#Handles clicking of links in help document
	$rtbHelp.add_LinkClicked({Invoke-Expression "start $($_.LinkText)"})
        
	$frmHelp.Controls.Add($rtbHelp)
	#endregion

	#region Buttons
	$BtnWidth = 80
	$BTT = @("Read Text","Exit")
	$Buttons = @(for($C=1;$C -le 2;$C++){New-Object -TypeName System.Windows.Forms.Button})
	$BC = $Buttons.Count
	for($C=0;$C -lt $Buttons.Count;$C++){
		$Buttons[$C].Anchor = Get-Anchor -B -R
		$Buttons[$C].Name = "Btn"+$BTT[$C]
		$Buttons[$C].Size = New-Object -TypeName System.Drawing.Size -ArgumentList $BtnWidth,30
		$Buttons[$C].Left = $rtbHelp.Right - ($BtnWidth*$BC)
		$Buttons[$C].Location = New-Object -TypeName System.Drawing.Point `
			-ArgumentList $Buttons[$C].Left,$($FrmHelp.Height*.86)
		$Buttons[$C].Text = $BTT[$C]
		$Buttons[$C].Enabled = $True
		$Buttons[$C].Visible = $True
		$Buttons[$C].UseVisualStyleBackColor = $True
		$FrmHelp.Controls.Add($Buttons[$C])
		$BC--
	}
	if(!$ReadText.IsPresent){
		$Buttons[0].Enabled = $False
		$Buttons[0].Visible = $False
	}
	$BtnRead_Click = {
		$FrmHelp.Cursor = Get-Cursor -Mode WaitCursor
		Add-Type -A “System.Speech”
		$SayIt = New-Object System.Speech.Synthesis.SpeechSynthesizer
		$SayIt.SelectVoiceByHints("Female")
		$SayIt.Speak($rtbHelp.Text)
		$FrmHelp.Cursor = Get-Cursor -Mode Default
	}
	$BtnExit_Click = {$FrmHelp.Close()}
	$Buttons[0].Add_Click($BtnRead_Click)
	$Buttons[1].Add_Click($BtnExit_Click)
	#endregion

	[Void]$frmHelp.ShowDialog()
}

<#
.NOTES
Name:    Invoke-OpenFileDialog Function
Author:  Randy Turner
Version: 1.0
Date:    03/22/2017

.SYNOPSIS
Provides a wrapper fumction used to Display an OpenFileDialog Control 
and return either the contentss of the selected file, the selected filename(s), 
or an empty string upon cancellation.

.PARAMETER ReturnMode. Alias: Mode
Optional, Used to specify the desired output: Contents(default)/Filename/Multiple

.PARAMETER Title  Alias: T
Optional, String used to set the OpenFileDialog.Title, defualt to $NULL.

.PARAMETER FileFilter  Alias: F
Optional, Filter String used to set the OpenFileDialog.Filter.
Defaults to "Text files (*.txt)|*.txt|All files (*.*)|*.*"

.PARAMETER FilterIndex  Alias: FI
One based index used to select the desired default file type

.PARAMETER InitialDirectory  Alias: Dir
used to set the Initial Directory of the control

.PARAMETER RestoreDirectory  Alias: RD
Determines wether the previously selected location is restored upon exit

.PARAMETER ShowHelp. Alias: ShowHelp
Determines if the Help button is shown, Overridden when
running under the 'Default Host' to True

.EXAMPLE
$FileContents = Invoke-OpenFileDialog
This example displays the OpenFileDialog and returns the selected file contents

.EXAMPLE
$Filename = Invoke-OpenFileDialog -ReturnMode Filename
This example displays the OpenFileDialog and returns the selected filename

.EXAMPLE
$Filename = Invoke-OpenFileDialog -ReturnMode Multiple
This example displays the OpenFileDialog and returns an array of selected filenames
#>

function Invoke-OpenFileDialog{
	param(
		[Parameter(Mandatory=$False)][Alias('Mode')]
			[ValidateNotNullOrEmpty()]
			[ValidateSet('Contents','Filename','Multiple')]
			[String]$ReturnMode = 'Contents',
		[Parameter(Mandatory=$False)][Alias('T')][String]$Title=$null,
		[Parameter(Mandatory=$False)][Alias('F')][String]$FileFilter='Text files (*.txt)|*.txt|All files (*.*)|*.*',
		[Parameter(Mandatory=$False)][Alias('FI')][Int]$FilterIndex=1,
		[Parameter(Mandatory=$False)][Alias('Dir')][String]$InitialDirectory=".",
		[Parameter(Mandatory=$False)][Alias('RD')][Switch]$RestoreDirectory,
		[Parameter(Mandatory=$False)][Alias('SH')][Switch]$ShowHelp)

	Add-Type -AssemblyName System.Windows.Forms

	<# 
	If Running under the Default Host $ShowHelp 
	MUST BE True or the Control Hangs.
	#>
	if($Host.Name -eq 'Default Host'){$ShowHelp = $True}

	$RV = ""
	$ReturnModes = @('Contents','Filename','Multiple')
	$ModeIdx = [Array]::IndexOf($ReturnModes,$ReturnMode)
	$OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.InitialDirectory = $InitialDirectory
	$OpenFileDialog.RestoreDirectory = $RestoreDirectory
	$OpenFileDialog.AutoUpgradeEnabled = $True
	$OpenFileDialog.Filter = $FileFilter
	$OpenFileDialog.FilterIndex = $FilterIndex
	$OpenFileDialog.Title = $Title
	$OpenFileDialog.Multiselect = ($ModeIdx -eq 2)
	$OpenFileDialog.ShowHelp = $ShowHelp
	if($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
		switch($ModeIdx){
			0 {$RV = Get-Content $OpenFileDialog.FileName}
			1 {$RV = $OpenFileDialog.FileName}
			2 {$RV = $OpenFileDialog.FileNames}
		}
	}
	return $RV
}