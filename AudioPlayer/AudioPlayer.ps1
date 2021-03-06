<#
.NOTES
-------------------------------------
Name:    AudioPlayer.ps1
Version: 1.0h - 02/17/2019
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------
Revision History:
This revision fixes an issue that
prevented files & Directories with 
paths containing square braces from
being loaded properly. A new option
was added  allowing the current
playlist to be reloaded directly.
-------------------------------------

.SYNOPSIS
This script launches Audio files from a definable location by use of a WinForm & Playlist.
--------------------------------------------------------------------------------------------
Supported Audio File Types are: .acc, .aif, .aiff, .au, .m4a, .mp3,.snd, .wav, .wma
--------------------------------------------------------------------------------------------

.DESCRIPTION

To Run the script without a console window use the included
PsRun3.exe PowerShell Launcher

PsRun3 accepts only 1 parameter - the name of the script to run and its parameters
enclosed in double quotes (") any quoted parameters should be enclosed in single 
quotes ('). You can create a shortcut with a parameter string like that below
to run the script for your location(s).

%PowerShellDevLib%\PsRun3.exe "%PowerShellDevLib%\AudioPlayer.ps1"
----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
---------------------------------------------------------------------------------------- 
.Parameter PlayList - Alias: Play
A text file with one entry per line naming a file or a directory of files to play.

.Parameter ErrorLog - Alias: Elog
Use to set the Error Log file, default: .\PS_Audio_Player_Errors.txt

.Parameter Volume - Alias: Vol
Used to set player volume, Values 0-100, default: -1 (use current)

.PARAMETER FontName - Alias: Fn
Name of the font to be used.

.PARAMETER FontSize - Alias: Fz
Size of the font to be used, between 9-24 points.

.PARAMETER FontStyle - Alias: Fs
Style of the font to be used: Bold, Italic, BoldItalic, Regular.

.PARAMETER Recurse - Alias: R
Use to recurse directories within the playlist.

.PARAMETER LoopPlayback - Alias: Loop
Use to Loop Playback of Playlist

.PARAMETER AutoPlay - Alias: Auto+
Use to Automatticly Play Command-Line Playlist on Load.

.PARAMETER AutoClose - Alias: Auto-
Use to Automattically Close the Player after completing Playlist.
Overrides -LoopPlayback

.PARAMETER LockVolume - Alias: LockV
Locks system audio volume

.PARAMETER HideLockVolume - Alias: LockH
Hides system audio volume lock

.PARAMETER LockSettings - Alias: LS
Used to Lock/Unlock registry settings, default: '' (Not Set)

.PARAMETER SaveSettings - Alias: Save
Used to Save registry settings

.EXAMPLE
PS> .\AudioPlayer.ps1 -Play "\\MediaServer01\Public\Shared Music\Playlists\NewOrleansJazz.apl" 
-Fn "Times New Roman" -Fz 16 -Fs Bold
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory = $False)][Alias('Play')][String]$PlayList = "",
	[Parameter(Mandatory = $False)][Alias('ELog')][String]$ErrorLog = ".\PS_Audio_Player_Errors.txt",
	[Parameter(Mandatory = $False)][Alias('Vol')]
		[ValidateNotNullOrEmpty()]
		[ValidateRange(0,100)]
		[Int]$Volume = -1, # -1 Indicates value not set
	[Parameter(Mandatory = $False)][Alias('Fn')][String]$FontName = "Lucida Console",
	[Parameter(Mandatory = $False)][Alias('Fz')]
		[ValidateNotNullOrEmpty()]
		[ValidateRange(9,24)]
		[Int]$FontSize = 12,
	[Parameter(Mandatory = $False)][Alias('Fs')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Bold','Italic','BoldItalic','Regular')]
		[String]$FontStyle = 'Regular',
	[Parameter(Mandatory = $False)][Alias('LS')]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Lock','Unlock')]
		[String]$LockSettings = '', # Indicates value not set
	[Parameter(Mandatory = $False)][Alias('R')][Switch]$Recurse,
	[Parameter(Mandatory = $False)][Alias('Loop')][Switch]$LoopPlayback,
	[Parameter(Mandatory = $False)][Alias('LockV')][Switch]$LockVolume,
	[Parameter(Mandatory = $False)][Alias('LockH')][Switch]$HideLockVolume,
	[Parameter(Mandatory = $False)][Alias('Auto+')][Switch]$AutoPlay,
	[Parameter(Mandatory = $False)][Alias('Auto-')][Switch]$AutoClose,
	[Parameter(Mandatory = $False)][Alias('Save')][Switch]$SaveSettings)

#Requires -Version 4 

#region Module Import
Import-Module -Name .\AppRegistry.ps1 -Force
Import-Module -Name .\Exists.ps1 -Force
Import-Module -Name .\ListViewSortLib.ps1 -Force
Import-Module -Name .\MetadataIndexLib.ps1 -Force
Import-Module -Name .\PCVolumeControl.ps1 -Force
Import-Module -Name .\WinFormsLibrary.ps1 -Force
#endregion

#region Script Level Variables
Add-Type -AssemblyName PresentationCore
$_MediaPlayer = New-Object -TypeName System.Windows.Media.MediaPlayer 
$Debug = $False
$PausePlayback = $True
$StopPlayback = $True
$LvwSortEnabled = $True
$PlayListErrors = $False 
$AppName = 'PS Audio Player'
$AppVers = 'Version: 1.0h - 02/17/2019'
$PlayListHeader = "*-<{0} - Playlist Header>-*" -f $AppName
$AudioFileTypes = @('.acc','.aif','.aiff','.au','.m4a','.mp3','.snd','.wav','.wma')
$HelpFont = @()
$RegKeys = @(
	'(Default)','Playlist','AutoClose','AutoPlay','HelpRtbFont',
	'HideVolumeLock','LockVolume','LoopPlayback','MainFormSize',
	'MainLvwColumnWidth','MainLvwFont','RecurseDirectory','Volume')
#endregion

#region Set Application Base Registry Key
$BaseKey = Get-RegistryProperty -App $AppName -Name $RegKeys[0]
if($BaseKey -ne $AppVers.Substring(9)){
	Set-RegistryProperty -App $AppName -Name $RegKeys[0] -Value $AppVers.Substring(9) -Type String}
#endregion

#region Get LockSettings value from Registry
# 3 Possible values: $Null(Non-Existent),0(Unlocked),1(Locked)
$LockSet = Get-RegistryProperty -App $AppName -Name 'LockSettings'
if($LockSet -eq $Null){$LockSet = 0}
# if $LockSettings set by user set new value
if($LockSettings.Length -gt 0){
	$LockSet = ConvertTo-Binary -Value $LockSettings
	Set-RegistryProperty -AppName $AppName -Name 'LockSettings' -Value $LockSet -Type Binary
	}
#endregion

#region Utility functions
function Get-SplitFilename{
	param(
		[Parameter(Mandatory = $True)][Alias('Fn')][String]$FileName,
		[Parameter(Mandatory = $False)][Alias('Eo')][Switch]$ExtensionOnly)
	if($ExtensionOnly -eq $False)
		{$RV = [IO.Path]::GetFileNameWithoutExtension($FileName)}
	else
		{$RV = [IO.Path]::GetExtension($FileName)}
	return $RV
}

function Get-SelectedFontStyle{
	param([Parameter(Mandatory = $True)][Alias('FS')][String]$SelectedFontStyle)

	$Styles = @("Bold","Italic","BoldItalic")

	switch([Array]::IndexOf($Styles,$SelectedFontStyle))
	{
		0	{[Drawing.FontStyle]::Bold}
		1	{[Drawing.FontStyle]::Italic}
		2	{[Drawing.FontStyle]([Drawing.FontStyle]::Bold -bor [Drawing.FontStyle]::Italic)}
		default {[Drawing.FontStyle]::Regular}
	}
}

function Get-ShortcutKey{
	param(
		[Parameter(Mandatory = $True)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('Open Playlist','New Playlist','Edit Playlist','Reload Playlist','Exit','Find','Find Next','Font Settings','Help','About','Lock Volume','Save Settings','Delete Settings')]
		[String]$Mode)

	$ValidModes = @('Open Playlist','New Playlist','Edit Playlist','Reload Playlist','Exit','Find','Find Next','Font Settings','Help','About','Lock Volume','Save Settings','Delete Settings')

	switch([Array]::IndexOf($ValidModes,$Mode))
		{
		 0 {[Windows.Forms.Keys]::Alt -bor [Windows.Forms.Keys]::O}
		 1 {[Windows.Forms.Keys]::Alt -bor [Windows.Forms.Keys]::N}
		 2 {[Windows.Forms.Keys]::Alt -bor [Windows.Forms.Keys]::E}
		 3 {[Windows.Forms.Keys]::Alt -bor [Windows.Forms.Keys]::R}
		 4 {[Windows.Forms.Keys]::Alt -bor [Windows.Forms.Keys]::F4}
		 5 {[Windows.Forms.Keys]::Alt -bor [Windows.Forms.Keys]::F}
		 6 {[Windows.Forms.Keys]::F3}
		 7 {[Windows.Forms.Keys]::Control -bor [Windows.Forms.Keys]::F}
		 8 {[Windows.Forms.Keys]::F1}
		 9 {[Windows.Forms.Keys]::Shift -bor [Windows.Forms.Keys]::F1}
		10 {[Windows.Forms.Keys]::Alt -bor [Windows.Forms.Keys]::L}
		11 {[Windows.Forms.Keys]::Control -bor [Windows.Forms.Keys]::S}
		12 {[Windows.Forms.Keys]::Control -bor [Windows.Forms.Keys]::Alt -bor [Windows.Forms.Keys]::D}
		}
}

function Load-Dll{
    param(
		[Parameter(Mandatory = $True)][Alias('DLL')][String]$FileName,
		[Parameter(Mandatory = $False)][Alias('Full')][Switch]$FullPath,
		[Parameter(Mandatory = $False)][Alias('Msg')][Switch]$ShowErrorMsg)

	if($FullPath.IsPresent){
		$DLLPath = $Filename}
	else{   
		$DLLPath = "{0}\{1}" -f '.',$FileName
	}
	if($ShowErrorMsg.IsPresent -and [IO.File]::Exists($DLLPath) -eq $False){
		$ErrMsg = "DLL File: {0} Missing or Invalid - Job Aborted!" -f $DLLPath
		$RV=Show-MessageBox -Msg $ErrMsg -Title $Form1.Text
		exit
	}
	Add-Type -Path $DLLPath
}

function Resolve-CurrentLocation{
	param([Parameter(Mandatory = $True)][Alias('I')][String]$Item)
	$D = Get-Location
	$F = Split-Path -Path $($Item) -Leaf
	return "{0}\{1}" -f $D,$F
}

function Invoke-Notepad{
	param([Parameter(Mandatory=$False)][ValidateNotNullOrEmpty()][ValidateSet("New","Edit")][String]$Mode="New")
	if($Mode -eq "New"){
		$F='.\New.apl'
		$PlayListHeader|Out-File -FilePath $F
		Notepad.exe $F
		Start-Sleep -Seconds 1
		Remove-Item -Path $F
	}
	else
		{Notepad.exe $Script:PlayList}
}

function Toggle-Boolean{
	param([Parameter(Mandatory = $True)][Bool]$Target)
	return ($Target = !$Target)
}

function Set-ButtonEnabledState{
	param(
		[Parameter(Mandatory = $True)]
		[ValidateNotNullOrEmpty()]
		[ValidateSet('PlayListLoading','PlayListLoaded','PlayClicked','PauseClicked','StopClicked','AllOff')]
		[String]$Mode = 'AllOff')
	$ValidModes = @('PlayListLoading','PlayListLoaded','PlayClicked','PauseClicked','StopClicked','AllOff')

#region ScriptBlocks
$CommonSB = {
	param([Boolean]$P1,[Boolean]$P2,[Boolean]$P3)
	$Buttons[0].Enabled = $P1
	$Buttons[1].Enabled = $P2
	$Buttons[2].Enabled = $P2
	0..0+3..5|ForEach-Object{$FileMenuItems[$_].Enabled = $P3}
	0..0+3..4+6|ForEach-Object{$LvwCtxMenuItems[$_].Enabled = $P3}
}
$PlayListLoading = {
	0..0+3..5|ForEach-Object{$FileMenuItems[$_].Enabled = !$FileMenuItems[$_].Enabled}
	0..0+3..4+6|ForEach-Object{$LvwCtxMenuItems[$_].Enabled = !$LvwCtxMenuItems[$_].Enabled}
}
$PauseClicked = {
	$Buttons[0].Enabled = $False
	$Buttons[2].Enabled = !$Buttons[2].Enabled
}
#endregion

	switch([Array]::IndexOf($ValidModes,$Mode)){
		0 {Invoke-Command -ScriptBlock $PlayListLoading}
		1 {Invoke-Command -ScriptBlock $CommonSB -ArgumentList (!$PlayListErrors),$False,$True}
		2 {Invoke-Command -ScriptBlock $CommonSB -ArgumentList $False,$True,$False}
		3 {Invoke-Command -ScriptBlock $PauseClicked}
		4 {Invoke-Command -ScriptBlock $CommonSB -ArgumentList $True,$False,$True}
		Default {foreach($Button in $Buttons){$Button.Enabled = $False}}}
}

function Get-RunTime{
	param([Parameter(Mandatory = $True)][Alias('P')][String]$Path)

	$ShellJob = {
		param([String]$Folder,[String]$File,$Descs,$Properties)
		#Create Windows Shell Object
		$Shell = New-Object -COMObject Shell.Application
		$ShellFolder = $Shell.Namespace($Folder)
		$ShellFile = $ShellFolder.ParseName($File)
		$RV = @()
		foreach($Index in $Properties)
			{$RV+=$ShellFolder.GetDetailsOf($ShellFile, $Index)}
		return $RV}

	$Props = $PropDescs = @("Length")
	# Get Property Index Numbers
	$MetaDataIndex = Get-MetadataIndex $path
	$SelectedProps = @()
	for($C=0;$C -le $Props.GetUpperBound(0);$C++)
		{$SelectedProps += Get-IndexByMetadataName $MetaDataIndex $Props[$C]}
	#Parse Path
	$Folder = Split-Path -Path $path
	$File = Split-Path -Path $path -Leaf
	$RV = Start-Job `
			-Name Get-RunTime `
			-ScriptBlock $ShellJob `
			-ArgumentList $Folder,$File,$PropDescs,$SelectedProps| `
		  Receive-Job -Wait -AutoRemoveJob
	return $RV
}

function Add-ListViewItem{
	param(
		[Parameter(Mandatory = $True)][Alias('Lvw')][Windows.Forms.ListView]$Control,
		[Parameter(Mandatory = $True)][Alias('Value')][String]$MainValue,
		[Parameter(Mandatory = $True)][Alias('Idx')][Int]$IconIndex)

	$Modulus = 2
	$FN = [IO.Path]::GetFileName($MainValue)
	$RunTime = Get-RunTime -Path $MainValue
	$LvwItem = New-Object -TypeName System.Windows.Forms.ListViewItem -ArgumentList $MainValue,$IconIndex
	[Void]$LvwItem.SubItems.Add($RunTime)
	[Void]$LvwItem.SubItems.Add($FN) 
	[Void]$Control.Items.Add($LvwItem)
	$LblStatus.Text = "Opening Playlist (File: {0}), Please Wait ..." -f $ListView[0].Items.Count
	if($Control.Items.Count % $Modulus -eq 0){[System.Windows.Forms.Application]::DoEvents()}
}

function Get-Children{
	param(
		[Parameter(Mandatory=$True)][String]$Parent,
		[Parameter(Mandatory=$False)][Alias('R')][Bool]$Recurse=$False)

	if($Recurse -eq $True)
		{Get-ChildItem -LiteralPath $Parent -Recurse}
	else
		{Get-ChildItem -LiteralPath $Parent}
}

function Set-FocusedListViewItem{
$ListView[0].Items[0].Selected = $True
$ListView[0].Items[0].EnsureVisible()
$ListView[0].Items[0].Focused = $True
}

function Validate-Playlist{
	param([Parameter(Mandatory = $True)][Alias('Items')][System.Array]$PlayListItems)

	$HdrValid = ($PlayListItems[0] -eq $PlayListHeader)
	$ItemExists = $False
	$ErrorList = @()
	if($HdrValid){
		$ArrayList = New-Object -TypeName System.Collections.ArrayList
		$ArrayList.AddRange($PlayListItems)
		$ArrayList.RemoveAt(0) #Remove Playlist Header
		$LineNo = 2 #1-Based Line# in Playlist
		foreach($Item in $ArrayList){
			if($Item.Length -eq 0 -or $Item.StartsWith('*')){$LineNo++;Continue}
			$FO = New-FileObject -Source $Item
			if($FO.IsDirectory)
				{$ItemExists = $FO.Path.Exists}
			else
				{$ItemExists = $FO.File.Exists}
			if(!$ItemExists){$ErrorList += $LineNo}
			$LineNo++
			}
		}
	$LineError = ($ErrorList.Length -gt 0)
	$NewProps = @{HdrIsValid = $HdrValid;LineErrors = $LineError;Array = $ArrayList;Errors = $ErrorList}
	$RV = New-Object -TypeName PSObject -Property $NewProps
	return $RV   
}

function Remove-BlankItems{
	param([Parameter(Mandatory = $True)][System.Collections.ArrayList]$Items)

	$BlankLines = @()
	for($C = 0;$C -lt $Items.Count;$C++){
		if($Items[$C].Length -eq 0 -or $Items[$C].StartsWith('*'))
			{$BlankLines += $C}
	}
	if($BlankLines.Length -gt 0){
		[Array]::Reverse($BlankLines)
		$BlankLines|ForEach-Object{$Items.RemoveAt($_)}
	}
	return $Items
}

function Report-PlaylistErrors{
	param([Parameter(Mandatory = $True)][Alias('Val')][PSObject]$Validation)

	$Script:PlayListErrors = $True
	if($AutoPlay.IsPresent)
		{# Output Error Log Header
		$AppName+" Errors {0}`r`nPlaylist: {1}`r`n" -f (Get-Date),$PlayList|Out-File -FilePath $ErrorLog
		if($Validation.HdrIsValid)
			{#Report Line Errors
			foreach($Line in $Validation.Errors)
				{"Line#: {0}" -f $Line|Out-File -FilePath $ErrorLog -Append}
			}
		else
			{"Invalid or Missing Playlist Header`r`n"|Out-File -FilePath $ErrorLog -Append}
		}
	else
		{
		if($Validation.HdrIsValid){
			$ErrMsg = "Playlist Errors Detected at Lines:`r`n{0}" `
				-f (Convert-ArrayToString -Array $RV.Errors)
			}
		else{
			$ErrMsg = "Playlist Errors Detected:`r`n"+
				"Invalid or Missing Playlist Header"
			}
		Show-MessageBox -Msg $ErrMsg -Title $Form1.Text
		}
}

function Load-ValidPlaylist{
# Vaild Items
$ListView[0].Items.Clear()
foreach($Line in $RV.Array){
	$CFO = New-FileObject -Source $Line 
	if($CFO.IsDirectory)
		{# Expand Directory to Child Files
		(Get-Children -Parent $Line $Recurse|Sort-Object -Property FullName).FullName| `
		ForEach-Object{
			if($AudioFileTypes.Contains([IO.Path]::GetExtension($_)))
				{Add-ListViewItem -Lvw $ListView[0] -Value $_ -Idx 1}}
		}
	else
		{Add-ListViewItem -Lvw $ListView[0] -Value $Line -Idx 1}
	}
	if($ListView[0].Items.Count -gt 0)
		{
		$FileMenuItems[3].Enabled = $True
		$LvwCtxMenuItems[3].Enabled = $True
		}
}

function Open-Playlist{
param([Parameter(Mandatory = $False)][Alias('NP')][Switch]$NoPrompt)

if($Script:LvwSortEnabled -eq $False)
	{
	$ErrMsg = "Open Playlist Disabled, During Current Operation"
	Show-MessageBox -Msg $ErrMsg -Title $Form1.Text
	}
else
	{
	$Script:LvwSortEnabled = Toggle-Boolean -Target $Script:LvwSortEnabled #Disable Sort

	$ListView[0].Columns[2].Width = -2 #Auto
	$FileFilter = $AppName+' Playlist (*.apl)|*.apl|Text files (*.txt)|*.txt|All files (*.*)|*.*'
    
	if(!$NoPrompt.IsPresent)
		{$RV = Invoke-OpenFileDialog -ReturnMode Filename -FileFilter $FileFilter -Title $AppName}
	else
		{$RV = $PlayList}

	if($RV.length -gt 0){
		$Script:PlayList = $RV
		$Script:PlayListErrors = $False
		if(Exists -Mode File $Script:PlayList){
			Set-ButtonEnabledState -Mode PlayListLoading
			$RV = Validate-Playlist -Items (Get-Content -Path $Script:PlayList)
			$RV.Array = Remove-BlankItems -Items $RV.Array
			if($RV.HdrIsValid -and $RV.LineErrors -eq $False)
				{Load-ValidPlaylist}
			else
				{Report-PlaylistErrors -Validation $RV}
			}
		}
	$LblStatus.Text = "Total Files in Playlist (After Expansion):  {0}" -f $ListView[0].Items.Count
	if($ListView[0].Items.Count -gt 0){
		Set-FocusedListViewItem
		Set-ButtonEnabledState -Mode PlayListLoaded}
	else{Set-ButtonEnabledState -Mode PlayListLoading}
	$ListView[0].Columns[2].Width -= ($ListView[0].Columns[2].Width*.03)
	$Script:LvwSortEnabled = Toggle-Boolean -Target $Script:LvwSortEnabled #Enable Sort
	}
}

function New-SearchValueItem{
New-Object -TypeName PSObject -Property @{Value = ''; Index = 0; Column = 0; Initialized = $False}}

function Find-ListViewItem{
	param(
		[Parameter(Mandatory = $True)][Alias('Lvw')][Windows.Forms.ListView]$LvwToSearch,
		[Parameter(Mandatory = $True)][Alias('Val')][String]$FindStr,
		[Parameter(Mandatory = $False)][Alias('Row')][Int]$StartPos=0,
		[Parameter(Mandatory = $False)][Alias('Col')][Int]$Column=0)
if($Script:LvwSortEnabled -eq $False)
	{
	$ErrMsg = "Find Disabled, During Current Operation"
	Show-MessageBox -Msg $ErrMsg -Title $Form1.Text
	}
else
	{
	$Script:LvwSortEnabled = Toggle-Boolean -Target $Script:LvwSortEnabled #Disable Sort
	for ($R = $StartPos; $R -lt $LvwToSearch.Items.Count; $R++){
		[System.Windows.Forms.ListViewItem]$Item=$LvwToSearch.Items[$R]
		if($($Item.SubItems[$Column].Text.ToLower()).Contains($FindStr.Tolower()) -eq $True)
			{
			$LvwToSearch.Items[$R].Selected = $True
			$LvwToSearch.SelectedItems[0].EnsureVisible() = $True
			$Script:FindObj.Index = $R++
			$Script:FindObj.Column = $Column
			$Script:FindObj.Value = $FindStr
			$Script:FindObj.Initialized = $True
			$R = $LvwToSearch.Items.Count
			}
		else
			{
			if($R -eq $LvwToSearch.Items.Count - 1){
				$Script:FindObj = New-SearchValueItem
				Show-MessageBox -Msg $("[{0}] Not Found!" -f $FindStr) -Title $Form1.Text}
			}
		}
	$Script:LvwSortEnabled = Toggle-Boolean -Target $Script:LvwSortEnabled #Enable Sort
	}
}

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

function Convert-ArrayToString{
	param(
		[Parameter(Mandatory = $True)][Array]$Array,
		[Parameter(Mandatory = $False)][String]$Delimiter=',')
	$FormatStr = ""
	for($C=0;$C -le $Array.GetUpperBound(0);$C++){$FormatStr += "{$C}$Delimiter"}
	return -join($FormatStr.SubString(0,$FormatStr.length - $Delimiter.length) -f $Array)
}

function New-IconCatalogItem{
	param(
		[Parameter(Mandatory = $False)][Alias('T')][String]$Tag='',
		[Parameter(Mandatory = $False)][Alias('C')][Int]$ControlIndex=0,
		[Parameter(Mandatory = $False)][Alias('I')][Int]$IconIndex=0)
	New-Object -TypeName PSObject -Property @{Tag = $Tag; ControlIndex = $ControlIndex; IconIndex = $IconIndex}
}

function New-FileObject{
	param([Parameter(Mandatory = $True)][String]$Source)
	$NewObj = New-Object -TypeName PSObject
	$File = New-Object -TypeName PSObject
	$Directory = New-Object -TypeName PSObject
	$Filename = [IO.Path]::GetFileName($Source)
	$Path = [IO.Path]::GetDirectoryName($Source)
	Add-Member -InputObject $Directory -MemberType NoteProperty -Name Exists -Value $([IO.Directory]::Exists($Path))
	Add-Member -InputObject $Directory -MemberType NoteProperty -Name Path -Value $Path
	Add-Member -InputObject $File -MemberType NoteProperty -Name Exists -Value $([IO.File]::Exists($Source))
	Add-Member -InputObject $File -MemberType NoteProperty -Name Name -Value $Filename
	Add-Member -InputObject $File -MemberType NoteProperty -Name NameWithoutExtension -Value $([IO.Path]::GetFileNameWithoutExtension($FileName))
	Add-Member -InputObject $File -MemberType NoteProperty -Name Extension -Value $([IO.Path]::GetExtension($Source))
	Add-Member -InputObject $NewObj -MemberType NoteProperty -Name FullPath -Value $([IO.Path]::GetFullPath($Source))
	Add-Member -InputObject $NewObj -MemberType NoteProperty -Name Path -Value $Directory
	Add-Member -InputObject $NewObj -MemberType NoteProperty -Name File -Value $File
	Add-Member -InputObject $NewObj -MemberType NoteProperty -Name IsDirectory -Value $($FileName.Length -eq 0)
return $NewObj
}

function Invoke-AudioFile{
	param(
		[Parameter(Mandatory = $True,ValueFromPipeline=$True)][Alias('F')][uri]$filepath,
		[Parameter(Mandatory = $False)][Alias('D')][Int]$OpenDelay = 2)
 
	$TimeFormat = "hh\:mm\:ss\.fff"
	$_MediaPlayer.Open($filepath)
	Start-Sleep -seconds $OpenDelay #This allows the player time to load the audio file
	$_MediaPlayer.Volume = 1
	$_MediaPlayer.Play()
	$LblDuration.Text = "Duration: [{0:$TimeFormat}]" -f $_MediaPlayer.NaturalDuration.TimeSpan
	Do	{
		$LblPosition.Text = "Position: [{0:$TimeFormat}]" -f $_MediaPlayer.Position
		if(([Audio]::Volume*100) -ne $Slider.Value){
			if($ToolMenuItems[1].Checked)
				#Enforce Volume Lock
				{[Audio]::Volume = $Slider.Value/100}
			else
				#Update Slider Value
				{$Slider.Value = [Int]([Audio]::Volume*100)}
			}
		[System.Windows.Forms.Application]::DoEvents()
		if($PausePlayback -eq $True){$_MediaPlayer.Pause()}
		}
	Until($_MediaPlayer.Position -eq $_MediaPlayer.NaturalDuration.TimeSpan -or $StopPlayback -eq $True)    
	$_MediaPlayer.Stop()
	$_MediaPlayer.Close()    
}

function Play-Playlist{
if($Script:LvwSortEnabled -eq $False){
	$ErrMsg = "Play Disabled, During Current Operation"
	Show-MessageBox -Msg $ErrMsg -Title $Form1.Text}
else
	{
	$Script:LvwSortEnabled = Toggle-Boolean -Target $Script:LvwSortEnabled #Disable Sort

	$Script:PausePlayback = $False
	$Script:StopPlayback = $False

	for($C=$ListView[0].SelectedItems[0].Index;$C -lt $ListView[0].Items.Count;$C++)
		{
		$ListView[0].Items[$C].Selected = $True
		$ListView[0].Items[$C].EnsureVisible()
		$LblStatus.Text = "Now Playing:  {0}" -f $ListView[0].SelectedItems[0].Subitems[2].Text
		Invoke-AudioFile -filepath $ListView[0].Items[$C].Text
		[System.Windows.Forms.Application]::DoEvents() 
		if($Script:StopPlayback -eq $True){$C=$ListView[0].Items.Count}
		if($Script:PausePlayback -eq $True){break}
		#Loopback Control
		if($CheckBoxes[0].Checked -and $C -eq $ListView[0].Items.Count-1){
			if($Script:AutoClose){
				Invoke-Command -ScriptBlock $Exit_Click
				break}
			$ListView[0].Items[0].Selected = $True
			$ListView[0].Items[0].EnsureVisible()
			$C = -1}
		}
	if($Script:AutoClose)
		{Invoke-Command -ScriptBlock $Exit_Click}
	else
		{Set-ButtonEnabledState -Mode StopClicked}
	$Script:LvwSortEnabled = Toggle-Boolean -Target $Script:LvwSortEnabled #Enable Sort
	}
}

function Save-Settings{
	if($LockSet -eq 1){
		Show-MessageBox -Msg 'Registry Settings Locked by Admin' -Title $AppName -S OkOnly -I Warning
		return
	}
	Set-RegistryProperty -App $AppName -Name $RegKeys[0] -Value $AppVers.Substring(9) -Type String
	Set-RegistryProperty -App $AppName -Name $RegKeys[1] -Value $PlayList -Type String
	Set-RegistryProperty -App $AppName -Name $RegKeys[2] -Value $CheckBoxes[1].Checked -Type Binary
	Set-RegistryProperty -App $AppName -Name $RegKeys[3] -Value $AutoPlay.IsPresent -Type Binary
	Set-RegistryProperty -App $AppName -Name $RegKeys[4] -Value $HelpFont -Type MultiString
	Set-RegistryProperty -App $AppName -Name $RegKeys[5] -Value $HideLockVolume.IsPresent -Type Binary
	Set-RegistryProperty -App $AppName -Name $RegKeys[6] -Value $ToolMenuItems[1].Checked -Type Binary
	Set-RegistryProperty -App $AppName -Name $RegKeys[7] -Value $CheckBoxes[0].Checked -Type Binary
	Set-RegistryProperty -App $AppName -Name $RegKeys[8] -Value @($Form1.ClientSize.Width,$Form1.ClientSize.Height) -Type MultiString
	$ToolMenuItems[3].Enabled = $True
	Set-RegistryProperty -App $AppName -Name $RegKeys[9] -Value @($ListView[0].Columns.Width) -Type MultiString
	Set-RegistryProperty -App $AppName -Name $RegKeys[10] -Value @($ListView[0].Font.Name,$ListView[0].Font.Size,$ListView[0].Font.Style) -Type MultiString
	Set-RegistryProperty -App $AppName -Name $RegKeys[11] -Value $CheckBoxes[1].Checked -Type Binary
	Set-RegistryProperty -App $AppName -Name $RegKeys[12] -Value $Slider.Value -Type DWord
}

function Get-Settings{
	$Values = New-Object -TypeName System.Collections.Generic.List[System.Object]
	$NewObj = New-Object -TypeName PSObject
	for($C=0;$C -lt $RegKeys.Count;$C++){
		$Values.Add((Get-RegistryProperty -App $AppName -Name $RegKeys[$C]))
		Add-Member -InputObject $NewObj -MemberType NoteProperty -Name $RegKeys[$C] -Value $Values[$C]
	}
	return $NewObj
}

function Load-Settings{
	param([Parameter(Mandatory = $True)][PSObject]$RS)

	$B2S = {param($B);[Switch]($B -eq 1)}

	if($Script:PlayList.Length -eq 0 -and $RS.Playlist.Length -gt 0){$Script:PlayList=$RS.Playlist}
	if(!$AutoPlay.IsPresent){$Script:AutoPlay = & $B2S $RS.AutoPlay}
	if(!$AutoClose.IsPresent){$Script:AutoClose = & $B2S $RS.AutoClose}
	if(!$Recurse.IsPresent){$Script:Recurse = & $B2S $RS.RecurseDirectory}
	if(!$LoopPlayback.IsPresent){$Script:LoopPlayback = & $B2S $RS.LoopPlayback}
	if(!$LockVolume.IsPresent){$Script:LockVolume = & $B2S $RS.LockVolume}
	if(!$HideLockVolume.IsPresent){$Script:HideLockVolume = & $B2S $RS.HideVolumeLock}
	if($Volume -eq -1 -and $RS.Volume -gt 0){$Script:Volume = $RS.Volume}
	if($RS.MainFormSize.Length -gt 0){
		Add-Member -InputObject $RS.MainFormSize -MemberType NoteProperty -Name Width  -Value $RS.MainFormSize[0]
		Add-Member -InputObject $RS.MainFormSize -MemberType NoteProperty -Name Height -Value $RS.MainFormSize[1]
	}
	if($RS.HelpRtbFont.Length -gt 0){
		$Script:HelpFont = $RS.HelpRtbFont
		$RS.HelpRtbFont = Parse-FontObject -ObjIn $RS.HelpRtbFont
	}
	if($RS.MainLvwFont.Length -gt 0){
		$RS.MainLvwFont = Parse-FontObject -ObjIn $RS.MainLvwFont
	}

	return $RS
}

function Parse-FontObject{
	param([Parameter(Mandatory = $True)][Object]$ObjIn)
	$Font = $ObjIn
	if($Font[2] -eq 'Bold, Italic'){$Font[2] ='BoldItalic'}
	$ObjIn = New-Object -TypeName PSObject
	Add-Member -InputObject $ObjIn -MemberType NoteProperty -Name Name  -Value $Font[0]
	Add-Member -InputObject $ObjIn -MemberType NoteProperty -Name Size  -Value $Font[1]
	Add-Member -InputObject $ObjIn -MemberType NoteProperty -Name Style -Value $Font[2]
	return $ObjIn
}
#endregion Utility functions

#region Add Custom DLL - Data Type: BlueflameDynamics.IconTools
Load-DLL -DLL "IconTools.dll" -Msg
#endregion

#region Script Level Variables
$FindObj = New-SearchValueItem
if($Playlist.StartsWith(".\")){$Playlist = Resolve-CurrentLocation $Playlist}
#endregion

$RegistrySettings = Load-Settings -RS (Get-Settings)

#region Build Icon DLL Catalog
<#	Array items defined in the order of the icons within the DLL
	The ControlIndex value of -1 identifies an unused icon, while
	other values set the desired sort order for importing the icons
	into 1 or more imagelists.
#>
$IconCatalog = @(
("App Icon","Open Playlist","New Playlist","Edit Playlist","Reload Playlist","Directory",
"Audio File","Find","Find Next","Font Settings","Help","Exit","Info"),
(-1,0,1,2,3,-1,-1,4,5,6,-1,7,-1)
)
$IconCatalogItems = @(for($C = 0; $C -lt $IconCatalog[0].Count; $C++)
	{New-IconCatalogItem -T $IconCatalog[0][$C] -I $C -C $IconCatalog[1][$C]})
# Filter down to selected members & Sort in ControlIndex Order 
$IconCatalogItems = $IconCatalogItems |`
Where-Object -Property ControlIndex -NotMatch -Value -1 |`
Sort-Object -Property ControlIndex
#endregion

function Show-MainForm{
	#region Import the Assemblies
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	Add-Type -AssemblyName Microsoft.VisualBasic
	#endregion

	#region Form Objects
	$Form1 = New-Object -TypeName Windows.Forms.Form
	$MainMenu = New-Object -TypeName Windows.Forms.MenuStrip
	$LvwCtxMenuStrip = New-Object -TypeName Windows.Forms.ContextMenuStrip
	$Slider = New-Object -TypeName Windows.Forms.TrackBar
	$Panel1 = New-Object -TypeName Windows.forms.Panel
	$InitialFormWindowState = New-Object -TypeName Windows.Forms.FormWindowState
	# Control Arrays
	$ListView = @(for ($C = 1; $C -le 1; $C++) {New-Object -TypeName Windows.Forms.ListView})
	$MainMenuItems = @(for ($C = 1; $C -le 3; $C++) {New-Object -TypeName Windows.Forms.ToolStripMenuItem})
	$FileMenuItems = @(for ($C = 1; $C -le 7; $C++) {New-Object -TypeName Windows.Forms.ToolStripMenuItem})
	$ToolMenuItems = @(for ($C = 1; $C -le 4; $C++) {New-Object -TypeName Windows.Forms.ToolStripMenuItem})
	$HelpMenuItems = @(for ($C = 1; $C -le 2; $C++) {New-Object -TypeName Windows.Forms.ToolStripMenuItem})
	$LvwCtxMenuItems = @(for ($C = 1; $C -le 8; $C++) {New-Object -TypeName Windows.Forms.ToolStripMenuItem})
	$LvwCtxMenuBars = @(for ($C = 1; $C -le 3; $C++) {New-Object -TypeName Windows.Forms.ToolStripSeparator})
	#endregion Form Objects

	#region ImageList for nodes
	$Script:ImageList = @(for ($C = 1; $C -le 2; $C++) {New-Object -TypeName Windows.Forms.ImageList})
	$Script:ImageList[0].ImageSize = New-Object -TypeName Drawing.Size -ArgumentList 24,24
	$Script:ImageList[1].ImageSize = New-Object -TypeName Drawing.Size -ArgumentList 24,24
	#endregion

	#region Custom Code for events.
	$About_Click = {
		$AboutText = -join(
		$AppVers,"{1}",
		"Created by Randy Turner - mailto:turner.randy21@yahoo.com","{2}",
		"PS Audio Player was designed as a specialized task launcher.","{2}",
		"Script Name: AudioPlayer.ps1","{2}",   
		"Synopsis:","{2}",
		"This script launches audio files from a defineable location by use of a WinForm","{1}",
		"{0}","{1}",
		"Supported Audio File Types: {3}","{1}",
		"{0}","{2}",
		"For additional help run the Powershell Get-Help cmdlet." `
		-f $("=" * 66),"`r`n",("`r`n"*2),(Convert-ArrayToString $AudioFileTypes))
		Show-AboutForm -AppName $AppName -AboutText $AboutText -URL
	}

	$Help_Click = {
	$HelpFile = ".\PS_Audio_Player_Help.txt"
	if($PausePlayback -or $StopPlayback){
		Show-HelpForm -AppName $AppName -HelpText $HelpFile -File -Read}
	else{
		Show-HelpForm -AppName $AppName -HelpText $HelpFile -File}
	}

	$Exit_Click = {$Form1.Close()}

	$FontSettings_Click = {Invoke-FontDialog -Control $ListView[0] -FontMustExist -AllowSimulations -AllowVectorFonts}

	$Find_Click = {
		$RV = Show-InputBox -Prompt "Find?" -Title "Search"
		if($RV -ne ""){Find-ListViewItem $ListView[0] $RV 0 2}  
	}

	$FindNext_Click = {
		if($Script:FindObj.Initialized -eq $True)
			{Find-ListViewItem $ListView[0] $Script:FindObj.Value $Script:FindObj.Index $Script:FindObj.Column}
		else
			{Invoke-Command -ScriptBlock $Find_Click}
	}

	$LockVolume_Click = {
		if($ToolMenuItems[1].Checked)
			{$ToolMenuItems[1].Text = 'Lock Volume'} #UnLock
		else
			{$ToolMenuItems[1].Text = 'Unlock Volume'} #Lock
		$Slider.Enabled = (!$Slider.Enabled)
		$ToolMenuItems[1].Checked = (!$ToolMenuItems[1].Checked)
	}

	$SaveSettings_Click = {Save-Settings}

	$DeleteSettings_Click = {
		if($LockSet -eq 1){
			Show-MessageBox -Msg 'Registry Settings Locked by Admim' -Title $AppName -S OkOnly -I Warning
			return
		}
		Flush-AppRegistryKey -App $AppName
		$This.Enabled = $False
	}

	$Form_Load_StateCorrection = {$Form1.WindowState = $InitialFormWindowState}
	#endregion

	#region Common Control Variables
	$DLL =  ".\AudioPlayerIcons.dll"
	<#Error handler for missing/invalid DLL#>
	if((Exists -Mode File -Location $DLL) -eq $False){
		$ErrMsg = "DLL File: {0} Missing or Invalid - Job Aborted!" -f $DLL
		$RV=Show-MessageBox -Msg $ErrMsg -Title $Form1.Text
		exit
	}
	$LvwCtxMenuItemsSize = New-Object -TypeName Drawing.Size -ArgumentList 219,32
	#endregion

	#region Form Level Code Groups
    #region Form Code
	$Form1.Text = $AppName
	$Form1.Name = "Form1"
	$Form1.FormBorderStyle = [Windows.Forms.FormBorderStyle]::Sizable
	$Form1.DataBindings.DefaultDataSourceUpdateMode = 0
	$Form1.Icon = [BlueflameDynamics.IconTools]::ExtractIcon($DLL,[Array]::IndexOf($IconCatalog[0],"App Icon"),16)
	$Form1.ClientSize = New-Object -TypeName Drawing.Size -ArgumentList 800,500
	$Form1.MinimumSize = New-Object -TypeName Drawing.Size -ArgumentList 620,412
	$Form1.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
	if($RegistrySettings.MainFormSize.Width -gt 0){
		$Form1.ClientSize = New-Object -TypeName Drawing.Size `
			-ArgumentList $RegistrySettings.MainFormSize.Width,$RegistrySettings.MainFormSize.Height
	}
	#endregion

	#region Populate Imagelists
	for ($C = 0; $C -le 1; $C++){$Script:ImageList[$C].Images.Clear()}
	[Drawing.Icon]$Ico = [BlueflameDynamics.IconTools]::ExtractIcon($DLL,[Array]::IndexOf($IconCatalog[0],"Directory"), 32)
	$Script:ImageList[0].Images.Add("Directory",$Ico)
	[Drawing.Icon]$Ico = [BlueflameDynamics.IconTools]::ExtractIcon($DLL,[Array]::IndexOf($IconCatalog[0],"Audio File"), 24)
	$Script:ImageList[0].Images.Add("Audio File",$Ico)
	for ($C = 0; $C -le $LvwCtxMenuitems.GetUpperBound(0); $C++){
		[Drawing.Icon]$Ico = [BlueflameDynamics.IconTools]::ExtractIcon($DLL, $IconCatalogItems[$C].IconIndex, 64)
		$Script:ImageList[1].Images.Add($IconCatalogItems[$C].Tag,$Ico)
	}
	#endregion 

	#region MainMenu 
	<#
	MainMenu is a Drop-Down menu designed to provide access to the
	various functions.
	#>
	$MainMenu.Name = "MainMenu"
	$MainMenu.Visible = $True
	$MainMenu.Size = New-Object -TypeName Drawing.Size -ArgumentList 220,30
	$MainMenu.Items.AddRange($MainMenuItems)
	$MainMenuItemsSize = New-Object -TypeName Drawing.Size -ArgumentList 219,22

	$FileMenuText = @("Open Playlist","New Playlist","Edit Playlist","Reload Playlist","Find","Find Next","Exit")
	Set-MenuItem $FileMenuText $FileMenuItems $MainMenuItemsSize "FileMenuItem"
	$FileMenuItems[0].Add_Click({Open-Playlist})
	$FileMenuItems[1].Add_Click({Invoke-Notepad -Mode New})
	$FileMenuItems[2].Add_Click({Invoke-Notepad -Mode Edit})
	$FileMenuItems[3].Add_Click({Open-Playlist -NP})
	$FileMenuItems[4].Add_Click($Find_Click)
	$FileMenuItems[5].Add_Click($FindNext_Click)
	$FileMenuItems[6].Add_Click($Exit_Click)
	$FileMenuItems[3].Enabled = $False
	$FileMenuBar = @(for ($C = 1; $C -le 2; $C++) {New-Object -TypeName Windows.Forms.ToolStripSeparator})

	$ToolMenuText = @('Font Settings','Lock Volume','Save Settings','Delete Settings')
	Set-MenuItem $ToolMenuText $ToolMenuItems $MainMenuItemsSize "ToolMenuItem" $False
	$ToolMenuItems[0].Add_Click($FontSettings_Click)
	$ToolMenuItems[1].Add_Click($LockVolume_Click)
	$ToolMenuItems[2].Add_Click($SaveSettings_Click)
	$ToolMenuItems[3].Add_Click($DeleteSettings_Click)
	$ToolMenuItems[1].Enabled = `
	$ToolMenuItems[1].Visible = !$HideLockVolume.IsPresent

	$HelpMenuText = @("Help","About")
	Set-MenuItem $HelpMenuText $HelpMenuItems $MainMenuItemsSize "HelpMenuItem"
	$HelpMenuItems[0].Add_Click($Help_Click)
	$HelpMenuItems[1].Add_Click($About_Click)

	$MainMenuText = @("File","Tools","Help")
	Set-MenuItem $MainMenuText $MainMenuItems $MainMenuItemsSize "MainMenuItem" $True
	$MainMenuItems[0].DropDownItems.AddRange($FileMenuItems)
	$MainMenuItems[1].DropDownItems.AddRange($ToolMenuItems)
	$MainMenuItems[2].DropDownItems.AddRange($HelpMenuItems)
	$MainMenuItems[0].DropDownItems.Insert(6,$FileMenuBar[1])
	$MainMenuItems[0].DropDownItems.Insert(4,$FileMenuBar[0])
	$Form1.Controls.Add($MainMenu)
	#endregion MainMenu

	#region Labels
	$LblDuration = New-Object -TypeName Windows.Forms.Label
	$LblDuration.Location = New-Object -TypeName Drawing.Point -ArgumentList 10,10
	$LblDuration.Name = "LblDuration"
	$LblDuration.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
	$LblDuration.BackColor = $CpColor = [System.Drawing.SystemColors]::Control
	$LblDuration.Text = ''
	$LblDuration.Width = 128
	$LblDuration.AutoEllipsis = $True
	$LblDuration.Anchor = Get-Anchor -T -L 
	$Panel1.Controls.Add($LblDuration)

	$LblPosition = New-Object -TypeName Windows.Forms.Label
	$LblPosition.Location = New-Object -TypeName Drawing.Point -ArgumentList 10,(10+$LblDuration.Height+5)
	$LblPosition.Name = "LblPosition"
	$LblPosition.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
	$LblPosition.BackColor = $CpColor
	$LblPosition.Text = ''
	$LblPosition.Width = $LblDuration.Width
	$LblPosition.AutoEllipsis = $True
	$LblPosition.Anchor = Get-Anchor -T -L
	$Panel1.Controls.Add($LblPosition)

	$LblStatus = New-Object -TypeName Windows.Forms.Label
	$LblStatus.Location = New-Object -TypeName Drawing.Point -ArgumentList (10+$LblDuration.Width+5),10
	$LblStatus.Name = "LblStatus"
	$LblStatus.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
	$LblStatus.BackColor = $CpColor
	$LblStatus.Text = ''
	$LblStatus.Width = $Panel1.Width - ($LblDuration.Width+20)
	$LblStatus.Height = $LblDuration.Height
	$LblStatus.AutoEllipsis = $True
	$LblStatus.Anchor = Get-Anchor -T -L -R
	$Panel1.Controls.Add($LblStatus)

	$LblVolume = New-Object -TypeName Windows.Forms.Label
	$LblVolume.Font = New-Object -TypeName System.Drawing.Font -ArgumentList "Microsoft Sans Serif",22,([Drawing.FontStyle]::Regular)
	$LblVolume.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 150,50
	$LblVolume.Name = "LblVolume"
	$LblVolume.BorderStyle = [Windows.Forms.BorderStyle]::None
	$LblVolume.Text = 'Vol: {0}' -f [Int]([Audio]::Volume*100)
	$LblVolume.Width = 125
	$LblVolume.Height = 45
	$LblVolume.AutoEllipsis = $True
	$LblVolume.Anchor = Get-Anchor -T -L
	$Panel1.Controls.Add($LblVolume)
	#endregion 

	#region Checkboxes
	$CBT = @('Loop Playback','Auto Close','Recurse Directories Opening Playlist')
	$CheckBoxes = @(for($C=1;$C -le 3;$C++){New-Object -TypeName System.Windows.Forms.Checkbox})
	for($C=0;$C -lt $CheckBoxes.Count;$C++){
		$CheckBoxes[$C].Anchor = Get-Anchor -T -L
		$CheckBoxes[$C].Width = 250
		$CheckBoxes[$C].Name = "Chk"+$($C+1)
		$CheckBoxes[$C].Location = New-Object -TypeName System.Drawing.Point `
			-ArgumentList 13,$($LblPosition.Bottom+(($CheckBoxes[$C].Height*$C)*.80))
		$CheckBoxes[$C].Text = $CBT[$C]
		$Panel1.Controls.Add($CheckBoxes[$C])
	}
	$CheckBoxes[0].Checked = $Script:LoopPlayback
	$CheckBoxes[1].Checked = $Script:AutoClose
	$CheckBoxes[2].Checked = $Script:Recurse
	$CheckBoxes[0].Add_CheckedChanged({$Script:LoopPlayback = !$Script:LoopPlayback})
	$CheckBoxes[1].Add_CheckedChanged({$Script:AutoClose = !$Script:AutoClose})
	$CheckBoxes[2].Add_CheckStateChanged({$Script:Recurse = !$Script:Recurse})
	#endregion

	#region LvwCtxMenuStrip
	$LvwCtxMenuStrip.Name = 'LvwCtxMenuStrip'
	$LvwCtxMenuStrip.Size = New-Object -TypeName Drawing.Size -ArgumentList 220,70
	$LvwCtxMenuStrip.Items.AddRange($LvwCtxMenuItems)
	#endregion

	#region LvwCtxMenuItems
	for($C = 0; $C -le $LvwCtxMenuItems.GetUpperBound(0); $C++){
		$LvwCtxMenuItems[$C].Name = 'ToolStripMenuItem' + $($C + 1)
		$LvwCtxMenuItems[$C].Text = $IconCatalogItems[$C].Tag
		$LvwCtxMenuItems[$C].Size = $LvwCtxMenuItemsSize
		$LvwCtxMenuItems[$C].Image = $Script:ImageList[1].Images[$C]
		$LvwCtxMenuItems[$C].ImageAlign = 'MiddleLeft'
		$LvwCtxMenuItems[$C].ShowShortcutKeys = $True
		$LvwCtxMenuItems[$C].ShortcutKeys = Get-ShortcutKey -Mode $LvwCtxMenuItems[$C].Text
	}
	$LvwCtxMenuItems[0].Add_Click({Open-Playlist})
	$LvwCtxMenuItems[1].Add_Click({Invoke-Notepad -Mode New})
	$LvwCtxMenuItems[2].Add_Click({Invoke-Notepad -Mode Edit})
	$LvwCtxMenuItems[3].Add_Click({Open-Playlist -NP})
	$LvwCtxMenuItems[4].Add_Click($Find_Click)
	$LvwCtxMenuItems[5].Add_Click($FindNext_Click)
	$LvwCtxMenuItems[6].Add_Click($FontSettings_Click)
	$LvwCtxMenuItems[7].Add_Click($Exit_Click)
	$LvwCtxMenuItems[3].Enabled = $False
	#endregion 

	#region Sepetator(s)
	$LvwCtxMenuStrip.Items.Insert(7, $LvwCtxMenuBars[2])
	$LvwCtxMenuStrip.Items.Insert(6, $LvwCtxMenuBars[1])
	$LvwCtxMenuStrip.Items.Insert(4, $LvwCtxMenuBars[0])
	#endregion

	#region listview[0] 
	$LvwColumnNames = @('','Duration','File')
	$LvwColumnWidths = @(28,-2,-2)
	$ListView[0].Name = "ListView0"
	$ListView[0].TabIndex = 0
	$ListView[0].View = [Windows.Forms.View]::Details
	$ListView[0].BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
	$ListView[0].MultiSelect = $False
	$ListView[0].GridLines = $True
	$ListView[0].FullRowSelect = $True
	$ListView[0].Size = New-Object -TypeName Drawing.Size -ArgumentList ($Form1.Width - 45),($Form1.Height - 250)
	$ListView[0].Location = New-Object -TypeName Drawing.Point -ArgumentList 13,28
	$ListView[0].TileSize = New-Object -TypeName Drawing.Size -ArgumentList $IconSize,$IconSize
	if($RegistrySettings.MainLvwFont -ne $Null){
		$FontName  = $RegistrySettings.MainLvwFont.Name
		$FontSize  = $RegistrySettings.MainLvwFont.Size
		$FontStyle = $RegistrySettings.MainLvwFont.Style
	}
	$ListView[0].Font = `
		New-Object -TypeName Drawing.Font -ArgumentList $FontName,$FontSize,(Get-SelectedFontStyle -FS $FontStyle)
	$ListView[0].SmallImageList = $imageList[0]
	$ListView[0].LargeImageList = $imageList[0]
	$ListView[0].ContextMenuStrip = $LvwCtxMenuStrip
	$ListView[0].UseCompatibleStateImageBehavior = $False
	$ListView[0].DataBindings.DefaultDataSourceUpdateMode = 0
	$ListView[0].Anchor = Get-Anchor -T -L -B -R
	for($C=0;$C -lt $LvwColumnNames.Count;$C++){
		$listView[0].Columns.Add($LvwColumnNames[$C])|Out-Null
		$listView[0].Columns[$C].Width=$LvwColumnWidths[$C]
	}
	$Form1.Controls.Add($ListView[0])
	$ColumnClick = {
		if($LvwSortEnabled -and $ListView[0].items.Count -gt 0){
			Sort-ListView -LvwControl $ListView[0] -Column $_.Column
			Set-FocusedListViewItem
		}
	}
	$ListView[0].Add_ColumnClick($ColumnClick)
	if($RegistrySettings.MainLvwColumnWidth -ne $Null){
		for($C=0;$C -lt $ListView[0].Columns.Count;$C++){
			$ListView[0].Columns[$C].Width = $RegistrySettings.MainLvwColumnWidth[$C]
		}
	}
	#endregion listvie[0] 

	#region Slider
	$CVol = if($Volume -eq -1){
				[Audio]::Volume*100}
			else{
				$Volume
				[Audio]::Volume=$Volume/100
				$LblVolume.Text = 'Vol: {0}' -f [Int]([Audio]::Volume*100)
				}
	$Slider.Name = 'VolumeSlider'
	$Slider.Visible = $True
	$Slider.Enabled = $True
	$Slider.TickStyle =[Windows.Forms.TickStyle]::Both
	$Slider.Text = 'Volume'
	$Slider.Width = 300
	$Slider.Minimum = 0
	$Slider.Maximum = 100
	$Slider.Value = $CVol
	$Slider.Location = New-Object -TypeName System.Drawing.Point -ArgumentList ($LblVolume.Right+5),40
	$Slider.BackColor= $CpColor
	$Slider.Anchor = Get-Anchor -T -L
	$Slider.Parent = $Panel1
	$Panel1.Controls.Add($Slider)
	$Slider.Add_ValueChanged({
		[Audio]::Volume = $This.Value/$This.Maximum
		$LblVolume.Text = 'Vol: {0}' -f $This.Value
		})
	#endregion

	#region Panel Control
	$Panel1.Location = New-Object -TypeName Drawing.Point -ArgumentList $ListView[0].Left,($ListView[0].bottom + 10)
	$Panel1.Size = New-Object -TypeName Drawing.Size -ArgumentList $ListView[0].width,130
	$Panel1.Visible = $True
	$Panel1.Parent = $Form1
	$Panel1.BorderStyle = [Windows.Forms.BorderStyle]::Fixed3D
	$Panel1.BackColor = [System.Drawing.Color]::LightGray
	$Panel1.Anchor = Get-Anchor -B -L -R
	$Form1.Controls.Add($Panel1)
	#endregion

	#region Buttons
	$BtnWidth = 80
	$BTT = @("Play","Pause","Stop")
	$Buttons = @(for($C=1;$C -le 3;$C++){New-Object -TypeName System.Windows.Forms.Button})
	$BC = $Buttons.Count
	for($C=0;$C -lt $Buttons.Count;$C++){
		$Buttons[$C].Anchor = Get-Anchor -B -R
		$Buttons[$C].Name = "Btn"+$BTT[$C]
		$Buttons[$C].Size = New-Object -TypeName System.Drawing.Size -ArgumentList $BtnWidth,30
		$Buttons[$C].Left = $Listview[0].Right - ($BtnWidth*$BC)
		$Buttons[$C].Location = `
			New-Object -TypeName System.Drawing.Point -ArgumentList $Buttons[$C].Left,$($Panel1.Bottom + 10)
		$Buttons[$C].Text = $BTT[$C]
		$Buttons[$C].Enabled = $False
		$Buttons[$C].UseVisualStyleBackColor = $True
		$Form1.Controls.Add($Buttons[$C])
		$BC--
	}
	$BtnPlay_Click = {
		Set-ButtonEnabledState -Mode PlayClicked;
		Play-Playlist}
	$BtnStop_Click = {
		Set-ButtonEnabledState -Mode StopClicked
		$Script:StopPlayback = $True
		$LblStatus.Text = ''}
	$BtnPause_Click = 
		{
		Set-ButtonEnabledState -Mode PauseClicked
		$Script:PausePlayback = !$Script:PausePlayback
		if($Script:PausePlayback -eq $True)
			{$Buttons[1].Text = 'Resume'}
		else
			{
			$Buttons[1].Text = 'Pause'
			$Script:_MediaPlayer.Play()
			}
		}
	$Buttons[0].Add_Click($BtnPlay_Click)
	$Buttons[1].Add_Click($BtnPause_Click)
	$Buttons[2].Add_Click($BtnStop_Click)
	#endregion
	#endregion Form Code

	#Save the initial state of the form
	$InitialFormWindowState = $Form1.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$Form1.Add_Load($Form_Load_StateCorrection)
	Set-ButtonEnabledState -Mode AllOff
	if($LockVolume -eq $True){Invoke-Command -ScriptBlock $LockVolume_Click}
	if($SaveSettings.IsPresent){Save-Settings}
	if($PlayList.Length -gt 0){
		$SplashImg = [BlueflameDynamics.IconTools]::ExtractIcon($DLL,[Array]::IndexOf($IconCatalog[0],"App Icon"),256)
		Import-Module -Name .\SplashScreen.ps1 -Force
		Show-SplashScreen -AppName $AppName -Image $SplashImg
		Open-Playlist -NoPrompt
		Close-SplashScreen
	}
	if($Script:AutoPlay.IsPresent){
		$LblStatus.Text = "AutoPlay Engaged, Please Wait ..."
		$DelayTimer = New-Object -TypeName System.Windows.Forms.Timer
		$DelayTimer.Enabled = $False
		$DelayTimeSeconds = 5
		$DelayTimer.Interval = 1000*$DelayTimeSeconds
		$DelayTimer_Tick={
			$DelayTimer.Stop()
			if($Script:PlayListErrors -eq $False)
				{Invoke-Command -ScriptBlock $BtnPlay_Click}
			elseif($Script:AutoClose -eq $True)
				{Invoke-Command -ScriptBlock $Exit_Click}                
		}
		$DelayTimer.Add_Tick($DelayTimer_Tick)
		$DelayTimer.Enabled = $True
		$DelayTimer.Start()
	}

	#Show the Form
	$Form1.BringToFront()
	[Void]$Form1.ShowDialog()    
}

#Call the Main function
Show-MainForm
#.\AudioPlayer.ps1 -PlayList .\Playlist.apl -Loop -Auto+ -Auto- -R -Vol 30 -LockV -LockH
#.\AudioPlayer.ps1 -PlayList .\Playlist_Errors.apl -Loop -Auto+ -Auto- -R