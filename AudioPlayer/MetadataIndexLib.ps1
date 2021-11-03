#Script Level Variables 
$path   = ""
$shell  = ""
$folder = ""
$file   = ""
$shellfolder = ""
$shellfile   = ""
$ShellObjInitialized = $False
$MaxPropertyIndex = 500

function Init-ShellObj
{
	param(
		[Parameter(Mandatory = $False)][Alias('P')][String]$TargetPath='.',
		[Parameter(Mandatory = $False)][Alias('M')][Int]$MPI=$Script:MaxPropertyIndex)
	#Create Windows Shell Object
	if($TargetPath){$Script:path = $TargetPath}
	$Script:shell = New-Object -COMObject Shell.Application
	$Script:folder = Split-Path -Path $Script:path
	$Script:file = Split-Path -Path $Script:path -Leaf
	$Script:shellfolder = $Script:shell.Namespace($Script:folder)
	$Script:shellfile = $Script:shellfolder.ParseName($Script:file)
	$Script:MaxPropertyIndex = $MPI
	$Script:ShellObjInitialized = $True
}

function Get-MetadataIndex
{
	param(
		[Parameter(Mandatory = $True)][Alias('P')][String]$TargetPath='',
		[Parameter(Mandatory = $False)][Alias('U')][Switch]$Unfiltered)
	#To get a list of index numbers and their named properties
	Init-ShellObj $TargetPath
	$MetadataIndex = 0..$Script:MaxPropertyIndex | Foreach-Object {New-MetadataIndexItem `
		-Index $_ `
		-Name  $Script:shellfolder.GetDetailsOf($null, $_)}
	if(!$Unfiltered){$MetadataIndex = $MetadataIndex|Where-Object -Property Name -ne -value ""}
	$MetadataIndex
}

function Show-MetadataIndex
{
	param(
		[Parameter(Mandatory = $True)][Alias('P')][String]$TargetPath='',
		[Parameter(Mandatory = $False)][Alias('U')][Switch]$Unfiltered)
	Get-MetadataIndex $TargetPath $Unfiltered
}

function New-MetadataIndexItem
{
	param(
		[Parameter(Mandatory = $False)][Alias('N')][String]$Name='',
		[Parameter(Mandatory = $False)][Alias('I')][Int]$Index=0)
	New-Object -TypeName PSObject -Property @{Name = $Name; Index = $Index}
}

function Get-IndexByMetadataName
{
	param(
		[Parameter(Mandatory = $True)][Alias('MDI')][Array]$MetaDataIndex,
		[Parameter(Mandatory = $True)][Alias('SV')][String]$SearchValue)
	$MetaDataIndex.Index[$MetaDataIndex.Name.IndexOf($SearchValue)]
}

function Get-ExtendedFileProperties
{
	param(
		[Parameter(Mandatory = $False)][Alias('P')][String]$Path,
		[Parameter(Mandatory = $False)][Alias('M')][Int]$MPI=$Script:MaxPropertyIndex)
	$Script:MaxPropertyIndex = $MPI
	Init-ShellObj $Path
	0..$MPI | Where-Object {$Script:shellfolder.GetDetailsOf($Script:shellfile, $_)} | 
		Foreach-Object{
			New-PropertyItem `
				-I $_ `
				-D $Script:shellfolder.GetDetailsOf($null, $_) `
				-V $Script:shellfolder.GetDetailsOf($Script:shellfile, $_)}
}

function New-PropertyItem
{
	param(
		[Parameter(Mandatory = $False)][Alias('I')][Int]$Index=0,
		[Parameter(Mandatory = $False)][Alias('D')][String]$Description="",
		[Parameter(Mandatory = $False)][Alias('V')][String]$Value="")
	New-Object -TypeName PSObject -Property @{Index = $Index;Description = $Description;Value=$Value}
}

<#
Import-Module .\MetadataIndexLib.ps1 -Force
Get-ExtendedFileProperties -P "c:\Users\turne\OneDrive\Documents\DVDFab11\FullDisc\Warehouse_13_S3E13.m4v"|Out-GridView
Get-MetadataIndex -P "\\MYBOOKLIVE\Public\Shared Videos\Movies\"|Out-GridView
#>