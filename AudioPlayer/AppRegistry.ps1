<#
.NOTES
Name:	AppRegistry.ps1
Author:  Randy Turner
Version: 1.0
Date:    04/01/2017

.SYNOPSIS
Provides a wrapper for utility fumctions used read/write registry entries.
#>

#region Script Level Variables
$DefaultBasePath = 'HKCU:\Software\BlueflameDynamics\'
$DefaultBinaryPairs = @('True','False','On','Off','Yes','No','Lock','Unlock','1','0')
#endregion

<#
.NOTES
Name:    Set-RegistryProperty Function
Author:  Randy Turner
Version: 1.0
Date:    04/01/2017

.SYNOPSIS
Provides a wrapper for fumction used to Create/Update a Registry Property

.PARAMETER BasePath Alias: Base
Optional, specifies Root Key Path

.PARAMETER AppName Alias: App
Required, Specifies the Application Key Name in the Root Path

.PARAMETER PropertyName Alias: Name
Required, Specifies the Property Name. 

.PARAMETER PropertyValue Alias: Value
Required, Specifies the Property Value to set.

.PARAMETER PropertyType Alias: Type
Required, Specifies the Property Value Data Type.

.EXAMPLE
Set-RegistryProperty -App $AppName -Name $Property -Value 30 -Type DWORD
This example creates/sets the $Appname Key property $Property to A DWORD value of 30
if the basepath, key, or property don't already exist they are created.
#>
function Set-RegistryProperty{
	param(
		[Parameter(Mandatory = $False)][Alias('Base')][String]$BasePath = $DefaultBasePath,
		[Parameter(Mandatory = $True)][Alias('App')][String]$AppName,
		[Parameter(Mandatory = $True)][Alias('Name')][String]$PropertyName,
		[Parameter(Mandatory = $True)][Alias('Value')][Object]$PropertyValue,
		[Parameter(Mandatory = $True)][Alias('Type')]
			[ValidateNotNullOrEmpty()]
			[ValidateSet('String','ExpandString','Binary','DWord','MultiString','QWord','Unknown')]
			[String]$PropertyType)
    
	$RegistryPath = -Join ($BasePath,$AppName,'\')

	#Create App Root, If Needed
	if(!(Test-Path -Path $RegistryPath)){New-Item -Path $RegistryPath -Force|Out-Null}
	#Add A Named Value
	New-ItemProperty `
		-Path $RegistryPath `
		-Name $PropertyName `
		-Value $PropertyValue `
		-PropertyType $PropertyType -Force|Out-Null
}

<#
.NOTES
Name:    Test-RegistryPropertyExists Function
Author:  Randy Turner
Version: 1.0
Date:    07/27/2018

.SYNOPSIS
Provides a wrapper for fumction used to test the existence of a Registry Property

.PARAMETER BasePath Alias: Base
Optional, specifies Root Key Path

.PARAMETER AppName Alias: App
Required, Specifies the Application Key Name in the Root Path

.PARAMETER PropertyName Alias: Name
Required, Specifies the Property Name. 

.EXAMPLE
Test-RegistryPropertyExists -App $AppName -Name $Property 
This example gets the $Appname Key property $Property value.
#>
function Test-RegistryPropertyExists{
	param(
		[Parameter(Mandatory = $False)][Alias('Base')][String]$BasePath = $DefaultBasePath,
		[Parameter(Mandatory = $True)][Alias('App')][String]$AppName,
		[Parameter(Mandatory = $True)][Alias('Name')][String]$PropertyName)
    
	$RegistryPath = -Join ($BasePath,$AppName,'\')
	$RV = Get-ItemProperty -Path $RegistryPath -Name $PropertyName -ErrorAction SilentlyContinue
    return $(if($RV){$True}else{$False})
}

<#
.NOTES
Name:    Get-RegistryProperty Function
Author:  Randy Turner
Version: 1.0
Date:    04/01/2017

.SYNOPSIS
Provides a wrapper for fumction used to get a Registry Property value

.PARAMETER BasePath Alias: Base
Optional, specifies Root Key Path

.PARAMETER AppName Alias: App
Required, Specifies the Application Key Name in the Root Path

.PARAMETER PropertyName Alias: Name
Required, Specifies the Property Name. 

.EXAMPLE
Get-RegistryProperty -App $AppName -Name $Property 
This example gets the $Appname Key property $Property value.
#>
function Get-RegistryProperty{
	param(
		[Parameter(Mandatory = $False)][Alias('Base')][String]$BasePath = $DefaultBasePath,
		[Parameter(Mandatory = $True)][Alias('App')][String]$AppName,
		[Parameter(Mandatory = $True)][Alias('Name')][String]$PropertyName)
    
	$RegistryPath = -Join ($BasePath,$AppName,'\')
	$PropExists = Get-ItemProperty -Path $RegistryPath -Name $PropertyName -ErrorAction SilentlyContinue
	if($PropExists){
		return (Get-ItemProperty -Path $RegistryPath -Name $PropertyName).$PropertyName}
	else {return $Null}
}

<#
.NOTES
Name:    Remove-RegistryProperty Function
Author:  Randy Turner
Version: 1.0
Date:    04/01/2017

.SYNOPSIS
Provides a wrapper for fumction used to Delete a Registry Property

.PARAMETER BasePath Alias: Base
Optional, specifies Root Key Path

.PARAMETER AppName Alias: App
Required, Specifies the Application Key Name in the Root Path

.PARAMETER PropertyName Alias: Name
Required, Specifies the Property Name. 

.EXAMPLE
Remove-RegistryProperty -App $AppName -Name $Property 
This example deletes the $Appname Key property $Property
#>
function Remove-RegistryProperty{
	param(
		[Parameter(Mandatory = $False)][Alias('Base')][String]$BasePath = $DefaultBasePath,
		[Parameter(Mandatory = $True)][Alias('App')][String]$AppName,
		[Parameter(Mandatory = $True)][Alias('Name')][String]$PropertyName)
    
	$RegistryPath = -Join ($BasePath,$AppName,'\')
	$PropExists = Get-ItemProperty -Path $RegistryPath -Name $PropertyName -ErrorAction SilentlyContinue
	if($PropExists){Remove-ItemProperty -Path $RegistryPath -Name $PropertyName}
}

<#
.NOTES
Name:    Flush-AppRegistryKey Function
Author:  Randy Turner
Version: 1.0
Date:    04/01/2017

.SYNOPSIS
Provides a wrapper for fumction used to Delete a Registry Key, its properties,
& the Application Base Key if the Application Key is the Last Application Key.
Used to remove all Application Properties in mass.

.PARAMETER BasePath Alias: Base
Optional, specifies Root Key Path

.PARAMETER AppName Alias: App
Required, Specifies the Application Key Name in the Root Path

.EXAMPLE
Flush-AppRegistryKey -App $AppName
This example gets the $Appname Key and associated properties in mass.
If the Application Key removed was the last on the basepath key it too is removed.
#>
function Flush-AppRegistryKey{
	param(
		[Parameter(Mandatory = $False)][Alias('Base')][String]$BasePath = $DefaultBasePath,
		[Parameter(Mandatory = $True)][Alias('App')][String]$AppName)
    
	$RegistryPath = -Join ($BasePath,$AppName,'\')
	if((Test-Path -Path $RegistryPath)){Remove-Item -Path $RegistryPath -Force|Out-Null}
	$Base = Get-ChildItem -Path $BasePath
	if($Base -eq $Null){Remove-Item -Path $BasePath -Force}
}

<#
.NOTES
Name:    ConvertTo-Binary Function
Author:  Randy Turner
Version: 1.0
Date:    04/01/2017

.SYNOPSIS
Provides a wrapper for fumction used to convert a String or Int32 to a Binary value.
when handling Int32 values, values greater than zero are True those less than 1 are False

.PARAMETER Value Alias: In
Required, specifies String or Int32 to be evaluated.

.PARAMETER BinaryPairs Alias: BP
Optional, An array of string pairs to be treated as True/False values.
All values are converted to lowercase before being evaluated.
Default is; @('True','False','On','Off','Yes','No','Lock','Unlock','1','0')
Even Index values (0,2,4,...) evaluate True.

.EXAMPLE
ConvertTo-Binary -In 'Yes'
This example converts the string 'Yes' to a value of 1.
Values returned are: 0(False), 1(True), or -1(Unrecognized Value)
#>
function ConvertTo-Binary{
	param(
		[Parameter(Mandatory = $True)][Alias('In')][Object]$Value,
		[Parameter(Mandatory = $False)][Alias('BP')]
			[System.Array]$BinaryPairs = $DefaultBinaryPairs)

	$Local:ErrMsg = 'Parameter: BinaryPairs Invalid! - Must be a System.Array Object with an even number of elements'
	if(($BinaryPairs.Length % 2) -ne 0){Throw $Local:ErrMsg}
	$BinaryPairs = $BinaryPairs | ForEach-Object { $_.ToLower() }
	$ObjType = $Value.GetType()
	if($ObjType.Name -eq 'String'){
		if($Value.Length -eq 0){return 0}
		$Value = $Value.ToLower()
		$ObjIdx = [Array]::IndexOf($BinaryPairs,$Value)
		if($ObjIdx -eq -1){
			return -1}
		else{
			return $(if(($ObjIdx % 2) -eq 0) {1} else {0})}
	}
	if($ObjType.Name -eq 'Int32'){
		return $(if($Value -lt 1){0}else{1})}
	return -1
}