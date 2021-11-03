<#Code snippet taken from:

https://www.milsuite.mil/book/docs/DOC-323444?ru=142390&sr=stream

#>




<###################################################################
  Active Directory Report Tool for USAF NIPR bases.

  NOTE 1: Edit the two Active Directory search paths to point to the
          relevant OUs for your base.

  NOTE 2: Script must be run from administrator account.

  NOTE 3: Each generated report is copied into system clipboard so 
          the data can be pasted into Excel or another application.

  20151111:  Initial script proof-of-concept.
  20151116:  Added options to check for empty groups and groups
             without managers.
  20151124:  Full GUI integration.
  20151125:  Added Admin Accounts button.
  20151126:  Added iPhone Users button.
  20151202:  Added auto copy to clipboard.
  20160126:  Added DMDC error check for user accounts.
  20160223:  Changed "ExpiredAccounts" to "KioskedAccounts".
  20161003:  Added disabled and locked users checks.
  20161102:  Re-worked User Count code based on ADUser-Numbers 
             script by Andrew Metzger.

  SHAUN CONRARDY, Contractor
  Cybersecurity, AF Systems
  shaun.conrardy.1.ctr@us.af.mil
###################################################################>

# ------------------------------------------------------------------
# Function for exit code on leaving application.
# ------------------------------------------------------------------
function OnApplicationExit {
	$script:ExitCode = 0 
}

# ------------------------------------------------------------------
# Primary function defining GUI and button functions.
# ------------------------------------------------------------------
function Call-ADTool_pff {

    # --------------------------------------------------------------
    # Set script version number.
    # --------------------------------------------------------------
    $ADTVersion = "2.20"

    # --------------------------------------------------------------
    # Set up Active Directory search paths.
    # Edit these to point to appropriate base.
    # --------------------------------------------------------------
	$SearchBase = "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
    $AdminSearchBase = "OU=Tyndall AFB,OU=Administrative Accounts,OU=Administration,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
    
    # --------------------------------------------------------------
    # Import system assemblies.
    # --------------------------------------------------------------
	[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")

    # --------------------------------------------------------------
    # Generate GUI form objects.
    # --------------------------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$formMain = New-Object System.Windows.Forms.Form
	$groupInfo = New-Object System.Windows.Forms.GroupBox
	$btnLocked = New-Object System.Windows.Forms.Button
	$btnDisabled = New-Object System.Windows.Forms.Button
	$btnDmdcError = New-Object System.Windows.Forms.Button
	$btnAdminAccounts = New-Object System.Windows.Forms.Button
	$btnKioskedAccounts = New-Object System.Windows.Forms.Button
	$btnPhones = New-Object System.Windows.Forms.Button
	$btnEmptyGroups = New-Object System.Windows.Forms.Button
	$btnUserCount = New-Object System.Windows.Forms.Button
	$btnInactiveUsers = New-Object System.Windows.Forms.Button
	$btnNoManagers = New-Object System.Windows.Forms.Button
	$lvMain = New-Object System.Windows.Forms.ListView
	$SB = New-Object System.Windows.Forms.StatusBar
	$menu = New-Object System.Windows.Forms.MenuStrip
	$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuFileExit = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
	$menuHelpAbout = New-Object System.Windows.Forms.ToolStripMenuItem
	$SBPStatus = New-Object System.Windows.Forms.StatusBarPanel
	$SBPBlog = New-Object System.Windows.Forms.StatusBarPanel
	$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
	
    # --------------------------------------------------------------
    # Create main GUI form window.
    # --------------------------------------------------------------
	$formMain_Load={
		[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
		$VBMsg = New-Object -COMObject WScript.Shell
        $Domain = ([DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).Name
		Set-FormTitle
	}
	
    # --------------------------------------------------------------
    # Handle clicking the "Exit" item under the "File" menu.
    # --------------------------------------------------------------
	$menuFileExit_Click={
		$formMain.Close()
	}
	
    # --------------------------------------------------------------
    # Handle clicking the "Empty Groups" button.
    # --------------------------------------------------------------
	$btnEmptyGroups_Click={
		Initialize-Listview
		$SBPStatus.Text = "Retrieving empty groups..."
		"Name","groupcategory" | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
        $Info = Get-ADGroup –Filter * -SearchBase $SearchBase -Properties Members | where { $_.Members.Count –eq 0 }
        Start-Sleep -m 400
        $Info | %{
            $Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
            foreach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){$Field = $Col.Text;$Item.SubItems.Add([string]$_.$Field)}
   	        $lvMain.Items.Add($Item)
        }
		$SBPStatus.Text = "Ready"
        
        # Copy data to system clipboard for pasting to Excel.
        $xText = "Empty Groups`tGroup Category`n"
        foreach($xitem in $lvMain.Items){
            foreach($xsubitem in $xitem.SubItems) {
                $xText += $xsubitem.Text+"`t"
            }
            $xText += "`n"
        }
        $xText | clip.exe
	}
	
    # --------------------------------------------------------------
    # Handle clicking the "No Managers" button.
    # --------------------------------------------------------------
	$btnNoManagers_Click={
		Initialize-Listview
		$SBPStatus.Text = "Retrieving groups with no manager..."
		"Name","groupcategory" | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
        $Info = Get-ADGroup -SearchBase $SearchBase -LdapFilter "(!(ManagedBy=*))"
        Start-Sleep -m 400
        $Info | %{
            $Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
            foreach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){$Field = $Col.Text;$Item.SubItems.Add([string]$_.$Field)}
   	        $lvMain.Items.Add($Item)
        }
		$SBPStatus.Text = "Ready"
        
        # Copy data to system clipboard for pasting to Excel.
        $xText = "Groups without Manager`tGroup Category`n"
        foreach($xitem in $lvMain.Items){
            foreach($xsubitem in $xitem.SubItems) {
                $xText += $xsubitem.Text+"`t"
            }
            $xText += "`n"
        }
        $xText | clip.exe
	}
	
    # --------------------------------------------------------------
    # Handle clicking the "User Counts" button.
    # --------------------------------------------------------------
	$btnUserCount_Click={
		Initialize-Listview
		$SBPStatus.Text = "Retrieving user counts..."
   		"Property","Count" | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
        
        # Gather the data from AD.
        $ALLusers = Get-ADUser -SearchBase $SearchBase -filter * -properties smartcardlogonrequired,employeetype

        # Set up variables and EmployeeType hash.
        $AllADUsers = @{}
        $employeeType = @()
        $hash = @{}
        $hash.A = "Active Duty Military"
        $hash.B = "Presidential Appointee"
        $hash.C = "DoD Civil Service"
        $hash.D = "Disabled American Veteran"
        $hash.E = "DoD Contractor"
        $hash.F = "Former Reserve Member - Retirement Eligible"
        $hash.G = "Organizational Account"
        $hash.H = "Medal of Honor Recipient"
        $hash.I = "Non-DoD Civil Service"
        $hash.J = "Academy Student"
        $hash.K = "Non-Appropriated Funds Employee"
        $hash.L = "Lighthouse Service"
        $hash.M = "Non-Government Agency Employee"
        $hash.N = "National Guard Member"
        $hash.O = "Non–DoD Contractor"
        $hash.Q = "Reserve Retiree - Not Eligible for Retired Pay"
        $hash.R = "Retired Military - Eligible for Retired Pay"
        $hash.S = "Resource Account"
        $hash.T = "Foreign Military"
        $hash.U = "Foreign National"
        $hash.V = "Reserve Member"
        $hash.NULL = "User with no EmployeeType"

        $employeetype += "NULL"
        $allAdusers.Null = @{}
        $AllAdusers.NULL.Accounts = @()
        $SCLnotrequired = @()
        $allnonorgbox = @()
        $disabled = @()

        foreach($entry in $allusers) {
	        if($employeetype -notcontains $entry.employeetype) {
		        $employeetype += $entry.employeetype
	        }
	        if($entry.employeetype -ne $NULL) {
		        $type = $entry.employeetype
		        if(!($AllADUsers."$type")) {
			        $allADUsers."$Type" = @{}
			        $AllAdusers."$type".Accounts = @()
		        }
	        }
	        if($entry.employeetype -eq $null) {$AllADUsers."NULL".Accounts += $entry}
	        else {$AllADUsers."$type".Accounts += $entry}
	        if($entry.employeetype -ne "G") {
	           if($entry.enabled -ne "True") {$disabled += $entry}
	           $allnonorgbox += $entry
	           if($entry.smartcardlogonrequired -ne $true) {$SCLnotrequired += $entry}
	        }
        }

        $allnonorgbox = $allusers | ?{$_.Employeetype -ne "G"}

        $employeetype = $employeetype | sort

        $totalcount = 0

        foreach($type in $employeetype) {
            $Item = New-Object System.Windows.Forms.ListViewItem($hash.$type)
            $Item.SubItems.Add([int]($allADusers."$type".Accounts.count))
            $lvMain.Items.Add($Item)
        }
        $Item2 = New-Object System.Windows.Forms.ListViewItem("Total User-Type Accounts")
        $Item2.BackColor = "LightGray"
        $Item2.SubItems.Add($allusers.count)
        $lvMain.Items.Add($Item2)
        $Item3 = New-Object System.Windows.Forms.ListViewItem("Total People Accounts")
        $Item3.BackColor = "LightGray"
        $Item3.SubItems.Add($allnonorgbox.count)
        $lvMain.Items.Add($Item3)
        $Item4 = New-Object System.Windows.Forms.ListViewItem("Disabled Accounts")
        $Item4.BackColor = "LightGray"
        $Item4.SubItems.Add($disabled.count)
        $lvMain.Items.Add($Item4)
        $Item5 = New-Object System.Windows.Forms.ListViewItem("SCL Not Required Accounts")
        $Item5.BackColor = "LightGray"
        $Item5.SubItems.Add($SCLnotrequired.count)
        $lvMain.Items.Add($Item5)

		$SBPStatus.Text = "Ready"
        
        # Copy data to system clipboard for pasting to Excel.
        $xText = "Property`tCount`n"
        foreach($xitem in $lvMain.Items){
            foreach($xsubitem in $xitem.SubItems) {
                $xText += $xsubitem.Text+"`t"
            }
            $xText += "`n"
        }
        $xText | clip.exe
	}
	
    # --------------------------------------------------------------
    # Handle clicking the "Inactive Users" button.
    # --------------------------------------------------------------
	$btnInactiveUsers_Click={
        $60days = (get-date).adddays(-60)
        $90Days = (get-date).adddays(-90)

		Initialize-Listview
		$SBPStatus.Text = "Retrieving inactive users..."
		"Name (yellow=60, red=90)" | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text

        $Info = Get-ADUser -SearchBase $SearchBase -filter * -properties lastlogontimestamp
        foreach($entry in $Info) {
            if(([datetime]::fromfiletime($entry.lastlogontimestamp)) -lt $90days) {
                $Item90 = New-Object System.Windows.Forms.ListViewItem($entry.Name)
                $Item90.BackColor = "Pink"
                $lvMain.Items.Add($Item90)
            }
            elseif(([datetime]::fromfiletime($entry.lastlogontimestamp)) -lt $60days) {
                $Item60 = New-Object System.Windows.Forms.ListViewItem($entry.Name)
                $Item60.BackColor = "Yellow"
                $lvMain.Items.Add($Item60)
            }
        }

		$SBPStatus.Text = "Ready"

        # Copy data to system clipboard for pasting to Excel.
        $xText = "Inactive Users`n"
        foreach($xitem in $lvMain.Items){
            foreach($xsubitem in $xitem.SubItems) {
                $xText += $xsubitem.Text+"`t"
            }
            $xText += "`n"
        }
        $xText | clip.exe
	}
	
    # --------------------------------------------------------------
    # Handle clicking the "Admin Accounts" button.
    # --------------------------------------------------------------
	$btnAdminAccounts_Click={
		Initialize-Listview
		$SBPStatus.Text = "Retrieving admin accounts..."
		"Name","LockedOut","Enabled" | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
        $Info = Get-ADUser -Filter * -Properties * -SearchBase $AdminSearchBase | Select-Object Name,LockedOut,Enabled | Sort-Object Name
        Start-Sleep -m 400
        $Info | %{
            $Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
            foreach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){
                $Field = $Col.Text
                $Item.SubItems.Add([string]$_.$Field)
                if ($Col.Text -eq "LockedOut" -AND $_.$Field -eq $true) {
			       $Item.BackColor = "Yellow"
		      	   $Item.ForeColor = "Black"
   	            }
                if ($Col.Text -eq "Enabled" -AND $_.$Field -eq $false) {
		  	       $Item.BackColor = "Yellow"
	   		       $Item.ForeColor = "Black"
       	        }
            }
            $lvMain.Items.Add($Item)
        }
		$SBPStatus.Text = "Ready"

        # Copy data to system clipboard for pasting to Excel.
        $xText = "Admin Account`tLockedOut?`tEnabled?`n"
        foreach($xitem in $lvMain.Items){
            foreach($xsubitem in $xitem.SubItems) {
                $xText += $xsubitem.Text+"`t"
            }
            $xText += "`n"
        }
        $xText | clip.exe
	}

    # --------------------------------------------------------------
    # Handle clicking the "iPhone Users" button.
    # --------------------------------------------------------------
	$btnPhones_Click={
		Initialize-Listview
		$SBPStatus.Text = "Retrieving iPhone users..."
		"Name" | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
        $Info = Get-ADGroup -Filter {Name -like 'GLS_*Good Mobile Users'} -SearchBase $SearchBase | Get-ADGroupMember | Select-Object Name | Sort-Object Name
        Start-Sleep -m 400
        $Info | %{
            $Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
            foreach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){$Field = $Col.Text;$Item.SubItems.Add([string]$_.$Field)}
   	        $lvMain.Items.Add($Item)
        }
		$SBPStatus.Text = "Ready"

        # Copy data to system clipboard for pasting to Excel.
        $xText = "iPhone Users`n"
        foreach($xitem in $lvMain.Items){
            foreach($xsubitem in $xitem.SubItems) {
                $xText += $xsubitem.Text+"`t"
            }
            $xText += "`n"
        }
        $xText | clip.exe
	}
	
    # --------------------------------------------------------------
    # Handle clicking the "Kiosked Accounts" button.
    # Still working out the kinks in this one.  Training date
    # doesn't show or isn't correct for all listed accounts.
    # --------------------------------------------------------------
	$btnKioskedAccounts_Click={
		Initialize-Listview
		$SBPStatus.Text = "Retrieving Kiosked Accounts..."
		"Name","iaTrainingDate" | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text

        $Info = Get-ADUser -Filter * -Properties * -SearchBase $SearchBase | Select-Object Name,MemberOf,iaTrainingDate | Sort-Object Name
        Start-Sleep -m 400
        $Info | %{
            [string]$GroupList = $_.MemberOf
            if ($GroupList -like "*GLS_ESD_IATRAINING_RESTRICTED*") {
                $Item = New-Object System.Windows.Forms.ListViewItem($_.$Col0)
                foreach ($Col in ($lvMain.Columns | ?{$_.Index -ne 0})){$Field = $Col.Text;$Item.SubItems.Add([string]$_.$Field)}
   	            $lvMain.Items.Add($Item)
            }
        }
		$SBPStatus.Text = "Ready"

        # Copy data to system clipboard for pasting to Excel.
        $xText = "Name`tIA Training Date`n"
        foreach($xitem in $lvMain.Items){
            foreach($xsubitem in $xitem.SubItems) {
                $xText += $xsubitem.Text+"`t"
            }
            $xText += "`n"
        }
        $xText | clip.exe
	}

    # --------------------------------------------------------------
    # Handle clicking the "DMDC Error" button.
    # --------------------------------------------------------------
	$btnDmdcError_Click={
		Initialize-Listview
		$SBPStatus.Text = "Retrieving DMDC error status on user accounts..."
		"User","DMDC Error","Enabled?" | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
        
        Get-ADUser -Filter 'info -like "*DMDC*"' -SearchBase $SearchBase -Properties * | foreach {
            $Item = New-Object System.Windows.Forms.ListViewItem($_.CN)
            $Item.SubItems.Add($_.info)
            $Item.SubItems.Add([string]$_.Enabled)
            $lvMain.Items.Add($Item)
        }

		$SBPStatus.Text = "Ready"
        
        # Copy data to system clipboard for pasting to Excel.
        $xText = "User`tDMDC Error`tEnabled`n"
        foreach($xitem in $lvMain.Items){
            foreach($xsubitem in $xitem.SubItems) {
                $xText += $xsubitem.Text+"`t"
            }
            $xText += "`n"
        }
        $xText | clip.exe
	}

    # --------------------------------------------------------------
    # Handle clicking the "Disabled Users" button.
    # --------------------------------------------------------------
	$btnDisabled_Click={
		Initialize-Listview
		$SBPStatus.Text = "Retrieving disabled user accounts..."
		"User","Enabled?" | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
        
        Get-ADUser -Filter {enabled -eq "false"} -SearchBase $SearchBase -Properties * | foreach {
            $Item = New-Object System.Windows.Forms.ListViewItem($_.CN)
            $Item.SubItems.Add([string]$_.Enabled)
            $lvMain.Items.Add($Item)
        }

		$SBPStatus.Text = "Ready"

        # Copy data to system clipboard for pasting to Excel.
        $xText = "User`tEnabled`n"
        foreach($xitem in $lvMain.Items){
            foreach($xsubitem in $xitem.SubItems) {
                $xText += $xsubitem.Text+"`t"
            }
            $xText += "`n"
        }
        $xText | clip.exe
	}

    # --------------------------------------------------------------
    # Handle clicking the "Locked Users" button.
    # --------------------------------------------------------------
	$btnLocked_Click={
		Initialize-Listview
		$SBPStatus.Text = "Retrieving locked user accounts..."
		"User","Locked?" | %{Add-Column $_}
		Resize-Columns
		$Col0 = $lvMain.Columns[0].Text
        
        Get-ADUser -Filter {locked -eq "true"} -SearchBase $SearchBase -Properties * | foreach {
            $Item = New-Object System.Windows.Forms.ListViewItem($_.CN)
            $Item.SubItems.Add([string]$_.LockedOut)
            $lvMain.Items.Add($Item)
        }

		$SBPStatus.Text = "Ready"

        # Copy data to system clipboard for pasting to Excel.
        $xText = "User`tLocked`n"
        foreach($xitem in $lvMain.Items){
            foreach($xsubitem in $xitem.SubItems) {
                $xText += $xsubitem.Text+"`t"
            }
            $xText += "`n"
        }
        $xText | clip.exe
	}

    # --------------------------------------------------------------
    # Handle clicking the "About" item under the "Help" menu.
    # --------------------------------------------------------------
	$menuHelpAbout_Click={
		Initialize-Listview
   		" "," " | %{Add-Column $_}
		Resize-Columns
		$Item = New-Object System.Windows.Forms.ListViewItem("Application Name")
        $Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
        $Item.Font = $Font
        $Item.BackColor = "LightGray"
	   	$Item.ForeColor = "Black"
        $Item.UseItemStyleForSubItems = $false
        $Item.SubItems.Add("Active Directory Report Tool")
		$lvMain.Items.Add($Item)
		$Item = New-Object System.Windows.Forms.ListViewItem("Application Version")
        $Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
        $Item.Font = $Font
        $Item.BackColor = "LightGray"
	   	$Item.ForeColor = "Black"
        $Item.UseItemStyleForSubItems = $false
		$Item.SubItems.Add($ADTVersion)
		$lvMain.Items.Add($Item)
		$Item = New-Object System.Windows.Forms.ListViewItem("Point of Contact")
        $Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
        $Item.Font = $Font
        $Item.BackColor = "LightGray"
	   	$Item.ForeColor = "Black"
        $Item.UseItemStyleForSubItems = $false
		$Item.SubItems.Add("See comment header in script.")
		$lvMain.Items.Add($Item)
		$SBPStatus.Text = "Ready"
	}
	
    # --------------------------------------------------------------
    # Function to get index of specified column name.
    # --------------------------------------------------------------
	function Get-ColumnIndex{
		Param($ColumnName)
		$Script:ColumnIndex = ($lvMain.Columns | ?{$_.Text -eq $ColumnName}).Index
	}
	
    # --------------------------------------------------------------
    # Function to reset/clear the display area of the window.
    # --------------------------------------------------------------
	function Initialize-Listview{
		$lvMain.Items.Clear()
		$lvMain.Columns.Clear()
	}
	
    # --------------------------------------------------------------
    # Function to add a column to the display area of the window.
    # --------------------------------------------------------------
	function Add-Column{
		Param([String]$Column)
		Write-Verbose "Adding $Column value"
		$lvMain.Columns.Add($Column)
	}
	
    # --------------------------------------------------------------
    # Function to resize columns equally across list view area.
    # --------------------------------------------------------------
	function Resize-Columns{
		Write-Verbose "Resizing columns based on column count"
		$ColWidth = (($lvMain.Width / ($lvMain.Columns).Count) - 11)
		$lvMain.Columns | %{$_.Width = $ColWidth}
	}
	
    # --------------------------------------------------------------
    # Function to remove GUI items from the form.
    # --------------------------------------------------------------
	function Remove-SelectedItems{
		$lvMain.SelectedItems | %{$lvMain.Items.RemoveAt($_.Index)}
	}
	
    # --------------------------------------------------------------
    # Function to set the window title.
    # --------------------------------------------------------------
	function Set-FormTitle{
		$formMain.Text = "AD Tool v$ADTVersion - Connected to " + $Domain	
	}
	
    # --------------------------------------------------------------
    # Correct initial state to prevent .Net maximized form issue.
    # --------------------------------------------------------------
	$Form_StateCorrection_Load={
		$formMain.WindowState = $InitialFormWindowState
	}
	
    # --------------------------------------------------------------
    # Remove all event handlers from the controls.
    # --------------------------------------------------------------
	$Form_Cleanup_FormClosed={
		try {
			$btnLocked.remove_Click($btnLocked_Click)
			$btnDisabled.remove_Click($btnDisabled_Click)
			$btnDmdcError.remove_Click($btnDmdcError_Click)
			$btnKioskedAccounts.remove_Click($btnKioskedAccounts_Click)
			$btnPhones.remove_Click($btnPhones_Click)
			$btnAdminAccounts.remove_Click($btnAdminAccounts_Click)
			$btnEmptyGroups.remove_Click($btnEmptyGroups_Click)
			$btnUserCount.remove_Click($btnUserCount_Click)
			$btnInactiveUsers.remove_Click($btnInactiveUsers_Click)
			$btnNoManagers.remove_Click($btnNoManagers_Click)
			$formMain.remove_Load($formMain_Load)
			$menuFileExit.remove_Click($menuFileExit_Click)
			$menuHelpAbout.remove_Click($menuHelpAbout_Click)
			$formMain.remove_Load($Form_StateCorrection_Load)
			$formMain.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch [Exception] { }
	}

    # --------------------------------------------------------------
    # GUI item definitions.
    # --------------------------------------------------------------
	#
	# Main Form
	#
	$formMain.Controls.Add($groupInfo)
	$formMain.Controls.Add($lvMain)
	$formMain.Controls.Add($SB)
	$formMain.Controls.Add($menu)
	$formMain.ClientSize = '780, 646'
	$formMain.MainMenuStrip = $menu
	$formMain.Name = "formMain"
	$formMain.StartPosition = 'CenterScreen'
	$formMain.Text = "Active Directory Report Tool v$ADTVersion"
	$formMain.add_Load($formMain_Load)
	#
	# Reports Button Grouping
	#
	$groupInfo.Controls.Add($btnLocked)
	$groupInfo.Controls.Add($btnDisabled)
	$groupInfo.Controls.Add($btnDmdcError)
	$groupInfo.Controls.Add($btnKioskedAccounts)
	$groupInfo.Controls.Add($btnPhones)
	$groupInfo.Controls.Add($btnAdminAccounts)
	$groupInfo.Controls.Add($btnEmptyGroups)
	$groupInfo.Controls.Add($btnUserCount)
	$groupInfo.Controls.Add($btnInactiveUsers)
	$groupInfo.Controls.Add($btnNoManagers)
	$groupInfo.Location = '10, 28'
	$groupInfo.Name = "groupInfo"
	$groupInfo.Size = '126, 336'
	$groupInfo.TabIndex = 7
	$groupInfo.TabStop = $False
	$groupInfo.Text = "Reports"
	#
	# Locked Button
	#
	$btnLocked.Location = '9, 298'
	$btnLocked.Name = "btnLocked"
	$btnLocked.Size = '110, 25'
	$btnLocked.TabIndex = 11
	$btnLocked.Text = "Locked Users"
	$btnLocked.UseVisualStyleBackColor = $True
	$btnLocked.add_Click($btnLocked_Click)
	#
	# Disabled Button
	#
	$btnDisabled.Location = '9, 267'
	$btnDisabled.Name = "btnDisabled"
	$btnDisabled.Size = '110, 25'
	$btnDisabled.TabIndex = 10
	$btnDisabled.Text = "Disabled Users"
	$btnDisabled.UseVisualStyleBackColor = $True
	$btnDisabled.add_Click($btnDisabled_Click)
	#
	# DMDC Error Button
	#
	$btnDmdcError.Location = '9, 236'
	$btnDmdcError.Name = "btnDmdcError"
	$btnDmdcError.Size = '110, 25'
	$btnDmdcError.TabIndex = 9
	$btnDmdcError.Text = "DMDC Error"
	$btnDmdcError.UseVisualStyleBackColor = $True
	$btnDmdcError.add_Click($btnDmdcError_Click)
	#
	# Kiosked Accounts Button
	#
	$btnKioskedAccounts.Location = '9, 205'
	$btnKioskedAccounts.Name = "btnKioskedAccounts"
	$btnKioskedAccounts.Size = '110, 25'
	$btnKioskedAccounts.TabIndex = 8
	$btnKioskedAccounts.Text = "Kiosked Accounts"
	$btnKioskedAccounts.UseVisualStyleBackColor = $True
	$btnKioskedAccounts.add_Click($btnKioskedAccounts_Click)
	#
	# iPhone Users Button
	#
	$btnPhones.Location = '9, 174'
	$btnPhones.Name = "btnPhones"
	$btnPhones.Size = '110, 25'
	$btnPhones.TabIndex = 7
	$btnPhones.Text = "iPhone Users"
	$btnPhones.UseVisualStyleBackColor = $True
	$btnPhones.add_Click($btnPhones_Click)
	#
	# Admin Account Button
	#
	$btnAdminAccounts.Location = '9, 143'
	$btnAdminAccounts.Name = "btnAdminAccounts"
	$btnAdminAccounts.Size = '110, 25'
	$btnAdminAccounts.TabIndex = 6
	$btnAdminAccounts.Text = "Admin Accounts"
	$btnAdminAccounts.UseVisualStyleBackColor = $True
	$btnAdminAccounts.add_Click($btnAdminAccounts_Click)
	#
	# Empty Groups Button
	#
	$btnEmptyGroups.Location = '9, 19'
	$btnEmptyGroups.Name = "btnEmptyGroups"
	$btnEmptyGroups.Size = '110, 25'
	$btnEmptyGroups.TabIndex = 2
	$btnEmptyGroups.Text = "Empty Groups"
	$btnEmptyGroups.UseVisualStyleBackColor = $True
	$btnEmptyGroups.add_Click($btnEmptyGroups_Click)
	#
	# User Count Button
	#
	$btnUserCount.Location = '9, 112'
	$btnUserCount.Name = "btnUserCount"
	$btnUserCount.Size = '110, 25'
	$btnUserCount.TabIndex = 5
	$btnUserCount.Text = "User Counts"
	$btnUserCount.UseVisualStyleBackColor = $True
	$btnUserCount.add_Click($btnUserCount_Click)
	#
	# Inactive Users Button
	#
	$btnInactiveUsers.Location = '9, 81'
	$btnInactiveUsers.Name = "btnInactiveUsers"
	$btnInactiveUsers.Size = '110, 25'
	$btnInactiveUsers.TabIndex = 4
	$btnInactiveUsers.Text = "Inactive Users"
	$btnInactiveUsers.UseVisualStyleBackColor = $True
	$btnInactiveUsers.add_Click($btnInactiveUsers_Click)
	#
	# Groups w/o Mgrs Button
	#
	$btnNoManagers.Location = '9, 50'
	$btnNoManagers.Name = "btnNoManagers"
	$btnNoManagers.Size = '110, 25'
	$btnNoManagers.TabIndex = 3
	$btnNoManagers.Text = "Groups w/o Mgr"
	$btnNoManagers.UseVisualStyleBackColor = $True
	$btnNoManagers.add_Click($btnNoManagers_Click)
	#
	# Main list view
	#
	$lvMain.Anchor = 'Top, Bottom, Left, Right'
	$lvMain.ContextMenuStrip = $contextMenu
	$lvMain.FullRowSelect = $True
	$lvMain.GridLines = $True
	$lvMain.Location = '142, 28'
	$lvMain.Name = "lvMain"
	$lvMain.Size = '630, 590'
	$lvMain.TabIndex = 13
	$lvMain.UseCompatibleStateImageBehavior = $False
	$lvMain.View = 'Details'
	#
	# Status Bar
	#
	$SB.Anchor = 'Bottom, Left, Right'
	$SB.Dock = 'None'
	$SB.Location = '0, 624'
	$SB.Name = "SB"
	[void]$SB.Panels.Add($SBPBlog)
	[void]$SB.Panels.Add($SBPStatus)
	$SB.ShowPanels = $True
	$SB.Size = '780, 22'
	$SB.TabIndex = 1
	$SB.Text = "Ready"
	#
	# Menu Bar
	#
	[void]$menu.Items.Add($menuFile)
	[void]$menu.Items.Add($menuHelp)
	$menu.Location = '0, 0'
	$menu.Name = "menu"
	$menu.Size = '780, 24'
	$menu.TabIndex = 0
	$menu.Text = "menuMain"
	#
	# File Menu
	#
	[void]$menuFile.DropDownItems.Add($menuFileExit)
	$menuFile.Name = "menuFile"
	$menuFile.Size = '37, 20'
	$menuFile.Text = "File"
	#
	# File Exit Menu Item
	#
	$menuFileExit.Name = "menuFileExit"
	$menuFileExit.Size = '186, 22'
	$menuFileExit.Text = "Exit"
	$menuFileExit.add_Click($menuFileExit_Click)
	#
	# Help Menu
	#
	[void]$menuHelp.DropDownItems.Add($menuHelpAbout)
	$menuHelp.Name = "menuHelp"
	$menuHelp.Size = '44, 20'
	$menuHelp.Text = "Help"
	#
	# Help About Menu Item
	#
	$menuHelpAbout.Name = "menuHelpAbout"
	$menuHelpAbout.Size = '152, 22'
	$menuHelpAbout.Text = "About"
	$menuHelpAbout.add_Click($menuHelpAbout_Click)
	#
	# Status Bar Text
	#
	$SBPStatus.AutoSize = 'Spring'
	$SBPStatus.Name = "Status"
	$SBPStatus.Text = "Ready"
	$SBPStatus.Width = 620
	#
	# Just some text in lower right corner
	#
	$SBPBlog.Alignment = 'Center'
	$SBPBlog.Name = "StatusLabel"
	$SBPBlog.Text = "AD Tool Status:"
	$SBPBlog.Width = 143
	# endregion Generated Form Code

    # --------------------------------------------------------------
    # Initialize form.
    # --------------------------------------------------------------
	# Save the initial state of the form
	$InitialFormWindowState = $formMain.WindowState
	# Init the OnLoad event to correct the initial state of the form
	$formMain.add_Load($Form_StateCorrection_Load)
	# Clean up the control events
	$formMain.add_FormClosed($Form_Cleanup_FormClosed)
	# Show the Form
	return $formMain.ShowDialog()
} 

# ------------------------------------------------------------------
# Main application area.
# ------------------------------------------------------------------
Import-Module ActiveDirectory
# Call the form
Call-ADTool_pff | Out-Null
# Perform cleanup
OnApplicationExit
