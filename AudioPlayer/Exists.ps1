<#
.NOTES
Name:    Exists.ps1
Author:  Randy Turner
Version: 2.0
Date:    06/20/2018

.SYNOPSIS
Provides a wrapper for fumctions used to test for the existance of a file or directory

.PARAMETER Mode
Required mode of operation FILE\DIRECTORY

.PARAMETER Location
File\Dirextory to validate.

.EXAMPLE
Exists -Mode File -Location "c:\Video\PF_Save_Summer.mp4"
This example returns True if the file exists.

.EXAMPLE
Exists -Mode Directory -Location "c:\Video\"
This example returns True if the directory exists.
#>
function Exists
{
	Param(
		[Parameter(Mandatory=$True)]
			[ValidateNotNullOrEmpty()]
			[ValidateSet("Directory","File")]
			[String]$Mode,
		[Parameter(Mandatory=$True)]
			[String]$Location)

	$Modes = @("Directory","File")

	Switch($Mode)
		{
		$Modes[0] {[IO.Directory]::Exists($Location)}
		$Modes[1] {[IO.File]::Exists($Location)}
		}
}

function Test-Exists
{
	Param(
		[Parameter(Mandatory=$True)]
			[ValidateNotNullOrEmpty()]
			[ValidateSet("Directory","File")]
			[String]$Mode,
		[Parameter(Mandatory=$True)]
			[String]$Location)

	$Modes = @("Directory","File")

	Switch($Mode)
		{
		$Modes[0] {[IO.Directory]::Exists($Location)}
		$Modes[1] {[IO.File]::Exists($Location)}
		}
}