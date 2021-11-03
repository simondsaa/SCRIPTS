<#
.NOTES
-------------------------------------
Name:    ListViewSortLib.ps1
Version: 1.0 - 04/01/2017
Author:  Randy E. Turner
Email:   turner.randy21@yahoo.com
-------------------------------------

.SYNOPSIS
This script contains functions required to Sort a ListView Control
in a Bi-Directional Manner by use of the ColumnClick Event. 
The functions are designed to support the use of multiple ListViews
----------------------------------------------------------------------------------------

.DESCRIPTION
The inital Click on a column header will sort the ListView items
in Ascending order, a second click will sort in Decending order, 
sucessive clicks toggle the sort order.

----------------------------------------------------------------------------------------
Security Note: This is an unsigned script, Powershell security may require you run the
Unblock-File cmdlet with the Fully qualified filename before you can run this script,
assuming PowerShell security is set to RemoteSigned.
---------------------------------------------------------------------------------------- 
#>

<#
.NOTES
Name:    Add-SortTrackers Function
Author:  Randy Turner
Version: 1.0
Date:    04/01/2017

.SYNOPSIS
This function adds one property group of three properties to a Winform ListView
These properties are used by the Sort-ListWiew Function to track information
used to control the sorting of a Listview based upon column clicks.
This function is normally called from the Sort-ListView Function.

.PARAMETER LvwControl
Required, the Target ListView Control

.OUTPUT
A property of SortTrackers is Added to the target ListView control consisting of 
3 sub-properties outlined below. 

LastColumnClicked tracks the last column number that was clicked
LastColumnAscending tracks the direction of the last sort of the active column
Initalized indicates if the Custom Properties Herein defined are Attached to the target control
#>
function Add-SortTrackers
{
	param([Parameter(Mandatory = $True)][Windows.Forms.ListView]$LvwControl)
	$Value = New-Object -TypeName PSObject -Property @{LastColumnClicked=-1;LastColumnAscending=$False;Initalized=$True}
	Add-Member -InputObject $LvwControl -MemberType NoteProperty -Name SortTrackers -Value $Value 
	return $LvwControl
}

<#
.NOTES
Name:    Sort-ListView Function
Author:  Randy Turner
Version: 1.0
Date:    04/01/2017

.SYNOPSIS
This function Serves as an Event Handler sorting the ListView in either 
Ascending/Descending order, The direction of a column is Ascending the 
first time a column header is clicked, a second successive click changes 
the sort to Descending, successive clicks toggle the direction of the sort. 
This function will call the Add-SortTrackers function the first time a 
listview column is sorted to attach custom properties to the ListView control 
used to control the sort process. This call occurs only if the Listview 
SortTrackers.Initialized property is null. To Add the event handler include 
the following replacing $ListView with the ListView Contol Name
$ListView.Add_ColumnClick({Sort-ListView -LvwControl $ListView -Column $_.Column})

.PARAMETER LvwControl
Required, the Target ListView Control

.PARAMETER Colomn
Required, the zero based ListView Column number passed during the ListView ColumnClick Event.
#>

function Sort-ListView
{
	param(
		[parameter(Position=0)][Windows.Forms.ListView]$LvwControl,
		[parameter(Position=1)][UInt32]$Column)

	if($LvwControl.SortTrackers.Initalized -eq $null){Add-SortTrackers -LvwControl $LvwControl}

	$Numeric = $true # determine direction how to sort
 
	<# 
	if the user clicked the same column that was clicked last time, reverse its sort order. 
	otherwise, reset for normal ascending sort
	#>
	if($LvwControl.SortTrackers.LastColumnClicked -eq $Column)
		{$LvwControl.SortTrackers.LastColumnAscending = !$LvwControl.SortTrackers.LastColumnAscending}
	else
		{$LvwControl.SortTrackers.LastColumnAscending = $true}

	$LvwControl.SortTrackers.LastColumnClicked = $Column

	$ListItems = @(@(@())) 
	<# 
	$ListItems is a three-dimensional array; 
	column 1 indexes the other columns, 
	column 2 is the value to be sorted on, and 
	column 3 is the System.Windows.Forms.ListViewItem object
	#>
 
	foreach($ListItem in $LvwControl.Items){
		# if all items are numeric, can use a numeric sort
		if($Numeric -ne $false) # nothing can set this back to true, so don't process unnecessarily
			{
			try   {$Test = [Double]$ListItem.SubItems[[int]$Column].Text}
			catch {$Numeric = $false} # a non-numeric item was found, so sort will occur as a string
			}
		$ListItems += ,@($ListItem.SubItems[[int]$Column].Text,$ListItem)}
 
	# create the expression that will be evaluated for sorting
	$EvalExpression = {
		if($Numeric)
			{return [Double]$_[0]}
		else
			{return [String]$_[0]}}
 
	# all information is gathered; perform the sort
	$ListItems = $ListItems | 
	Sort-Object -Property `
	@{Expression=$EvalExpression; Ascending=$LvwControl.SortTrackers.LastColumnAscending}
 
	## the list is sorted; display it in the listview
	$LvwControl.BeginUpdate()
	$LvwControl.Items.Clear()
	foreach($ListItem in $ListItems){$LvwControl.Items.Add($ListItem[1])}
	$LvwControl.EndUpdate()
}