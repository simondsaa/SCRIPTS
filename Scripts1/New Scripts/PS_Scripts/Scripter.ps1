# Written by SSgt Timothy Brady
# Tyndall AFB, Panama City, FL
# Created February 5, 2016

# Function Start
# ----------------------------------------------------------------------------
Function GenerateForm 
{
    # Loading Required Assemblies
    # ----------------------------------------------------------------------------
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
    [Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null
    
    # Setting all the form tools to variables
    # ----------------------------------------------------------------------------
    $form1 = New-Object System.Windows.Forms.Form
    $button1 = New-Object System.Windows.Forms.Button
    $button2 = New-Object System.Windows.Forms.Button
    $textBox1 = New-Object System.Windows.Forms.TextBox
    $ListView = New-Object System.Windows.Forms.ListView
    $label1 = New-Object System.Windows.Forms.Label
    $label2 = New-Object System.Windows.Forms.Label
    $checkBox7 = New-Object System.Windows.Forms.CheckBox
    $checkBox6 = New-Object System.Windows.Forms.CheckBox
    $checkBox5 = New-Object System.Windows.Forms.CheckBox
    $checkBox4 = New-Object System.Windows.Forms.CheckBox
    $checkBox3 = New-Object System.Windows.Forms.CheckBox
    $checkBox2 = New-Object System.Windows.Forms.CheckBox
    $checkBox1 = New-Object System.Windows.Forms.CheckBox
 
    # "Run" Button Action
    # ----------------------------------------------------------------------------
    $button1Click=
    {
        # Check Box 1 Script
        # ----------------------------------------------------------------------------
        If ($checkBox1.Checked)
        {
            $ListView.Items.Add($checkBox1.Text).BackColor = "Silver"
            
            $Computer = $textBox1.Text
            
            If (!($Computer -eq ""))
            {
                If (Test-Connection $Computer -Count 1 -ea 0)
                {
                    Try
	                {        
	                    $RegOpen = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer)
	                    $RegKey = $RegOpen.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation')
	                    $SDC = $RegKey.GetValue('Model')
	                }
	                Catch
	                {
	                    $SDC = "N/A"
	                }
                    
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
                    
                    $ListView.Items.Add("System Name        : $SysName")
                    $ListView.Items.Add("Operating System   : $OS $SP")
                    $ListView.Items.Add("SDC Version        : $SDC")
                    $ListView.Items.Add("System Bit         : $Bit")
                    $ListView.Items.Add("Processor          : $CPU")
                    $ListView.Items.Add("Physical Memory    : $RAM MB")
                    $ListView.Items.Add("Manufacturer       : $Manufacturer")
                    $ListView.Items.Add("Model              : $Model")
                    $ListView.Items.Add("Serial Number      : $Serial")
                    $ListView.Items.Add("IP Address         : $IP")
                    $ListView.Items.Add("MAC Address        : $MAC")
                    $ListView.Items.Add("System Uptime      : $UpDays day(s) $UpHours hours $UpMins mins")
                }
            }
            Else
            {
                $ListView.Items.Add("No computer name entered...")
            }
        }
 
        # Check Box 2 Script
        # ----------------------------------------------------------------------------
        If ($checkBox2.Checked)
        {
            $ListView.Items.Add($checkBox2.Text).BackColor = "Silver"
            
            $Computer = $textBox1.Text
            
            If (!($Computer -eq ""))
            {
                If (Test-Connection $Computer -Count 1 -ea 0)
                {
                    $User = Get-WmiObject Win32_ComputerSystem -cn $Computer -ErrorAction SilentlyContinue
                    If ($User.UserName -ne $null)
                    {
                        $UserName = $User.UserName

                        $ListView.Items.Add("$Computer  - Current logged on user - $UserName")
                        
                        $EDI = $User.UserName.Split("\")[1]
                        $UserInfo = Get-ADUser "$EDI" -Properties DisplayName, City, EmailAddress, extensionAttribute5, mDBOverHardQuotaLimit, LockedOut, Enabled, OfficePhone -ErrorAction SilentlyContinue
                        
                        $MailSize = ($UserInfo.mDBOverHardQuotaLimit/1024)
                        $OU = ($UserInfo.distinguishedName -split ",OU=")[1]
                        $Name = $UserInfo.DisplayName
                        $Pre = $UserInfo.SamAccountName
                        $Base = $UserInfo.City
                        $Email = $UserInfo.EmailAddress
                        $Cat = $UserInfo.extensionAttribute5
                        $Locked = $UserInfo.LockedOut
                        $Enabled = $UserInfo.Enabled
                        $Number = $UserInfo.OfficePhone
                        
                        $ListView.Items.Add("Display Name       : $Name")
                        $ListView.Items.Add("Pre-Windows 2000   : $Pre")
                        $ListView.Items.Add("Base Name          : $Base")
                        $ListView.Items.Add("Email Address      : $Email")
                        $ListView.Items.Add("Email Category     : $Cat")
                        $ListView.Items.Add("Mailbox Size Limit : $MailSize MB")
                        $ListView.Items.Add("Account Locked Out : $Locked")
                        $ListView.Items.Add("Account Enabled    : $Enabled")
                        $ListView.Items.Add("Phone Number       : $Number")
                    }

                    Else 
                    {
                        $UserName = "No User Logged On"
                        $ListView.Items.Add("$Computer      : Current logged on user - $UserName")
                    }
                    
                }
                
                Else
                {
                    $ListView.Items.Add("$Computer is either an invalid name or is offline")
                }
            }
            Else
            {
                $ListView.Items.Add("No computer name entered...")
            }
        }
 
        # Check Box 3 Script
        # ----------------------------------------------------------------------------
        If ($checkBox3.Checked)
        {
            $ListView.Items.Add($checkBox3.Text).BackColor = "Silver"

            $Computer = $textBox1.Text
            If (!($Computer -eq ""))
            {
                $PsTools = Get-ChildItem -Path "\\xlwu-fs-05pv\Tyndall_PUBLIC\Ncc Admin\Tools\PsTools"
                ForEach ($Item in $PsTools)
                {
                    Copy-Item $Item.FullName -Destination "C:\Windows\System32" -Force
                }
                
                Start-Sleep -Seconds 3
                cd C:\Windows\System32
                psexec \\$Computer /accepteula -d c:\windows\system32\winrm.cmd quickconfig -quiet | Out-Null
            }
            Else
            {
                $ListView.Items.Add("No computer name entered...")
            }
        }

        # Check Box 4 Script
        # ----------------------------------------------------------------------------
        If ($checkBox4.Checked)
        {
            $ListView.Items.Add($checkBox4.Text).BackColor = "Silver"

            $Computer = $textBox1.Text
            If (!($Computer -eq ""))
            {
                $PsTools = Get-ChildItem -Path "\\xlwu-fs-05pv\Tyndall_PUBLIC\Ncc Admin\Tools\PsTools"
                ForEach ($Item in $PsTools)
                {
                    Copy-Item $Item.FullName -Destination "C:\Windows\System32" -Force
                }

                # Loads the file selection window
                # ----------------------------------------------------------------------------
                [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
                $dialog = New-Object System.Windows.Forms.OpenFileDialog
                $dialog = New-Object System.Windows.Forms.OpenFileDialog
                $dialog.FilterIndex = 0
                $dialog.InitialDirectory = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Applications"
                $dialog.Multiselect = $false
                $dialog.RestoreDirectory = $true
                $dialog.Title = "Select File"
                $dialog.ValidateNames = $true
                $dialog.ShowDialog()

                #$ListView.Items.Add("Running...")

                cd C:\Windows\System32
                psexec \\$Computer /accepteula -s -i -d $dialog.FileName | Out-Null

                $ListView.Items.Add("Finished").BackColor = "LightGreen"
            }
            Else
            {
                $ListView.Items.Add("No computer name entered...")
            }
        }

        # Check Box 5 Script
        # ----------------------------------------------------------------------------
        If ($checkBox5.Checked)
        {
            $ListView.Items.Add($checkBox5.Text).BackColor = "Silver"
            
            Start-Process cmd.exe
        }
        
        # Check Box 6 Script
        # ----------------------------------------------------------------------------
        If ($checkBox6.Checked)
        {
            $ListView.Items.Add($checkBox6.Text).BackColor = "Silver"
            
            Start-Process Powershell_ISE.exe
        }
        
        # Check Box 7 Script
        # ----------------------------------------------------------------------------
        If ($checkBox7.Checked)
        {
            $ListView.Items.Add($checkBox7.Text).BackColor = "Silver"
            
            $User = Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue
            $LocalUser = $User.UserName.Split("\")[1]
            $Documents = "C:\Users\$LocalUser\Documents"
                        
            
            # Loads the file selection window
            # ----------------------------------------------------------------------------
            [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Filter = 'PowerShell Files|*.ps1|All Files|*.*'
            $dialog.FilterIndex = 0
            $dialog.InitialDirectory = $Documents
            $dialog.Multiselect = $false
            $dialog.RestoreDirectory = $true
            $dialog.Title = "Select File"
            $dialog.ValidateNames = $true
            $dialog.ShowDialog()

            $Script = $dialog.FileName

            Start-Process PowerShell.exe -ArgumentList "-file `"$Script`""
        }
 
        If (!$checkBox1.Checked -and !$checkBox2.Checked -and !$checkBox3.Checked -and !$checkBox4.Checked -and !$checkBox5.Checked -and !$checkBox6.Checked -and !$checkBox7.Checked)
        {
            $ListView.Items.Add("No CheckBox selected....")
        }
    }

    # "Clear Host" Button Action
    # ----------------------------------------------------------------------------
    $button2Click=
    {
        $ListView.Items.Clear()
    }

    $OnLoadForm_StateCorrection=
    {
        $form1.WindowState = $InitialFormWindowState
    }
 
# Generating the Form or Scripter Window
# ----------------------------------------------------------------------------
    
    # Form 1
    
    $form1.Text = "Script Launcher"
    $form1.Name = "form1"
    $form1.Width = 1000
    $form1.Height = 400
    $form1.FormBorderStyle = "Fixed3D"
    $form1.MaximizeBox = $false
    $form1.MinimizeBox = $true

    # Button 1
    
    $button1.Name = "button1"
    $button1.Width = 75
    $button1.Height = 25
    $button1.UseVisualStyleBackColor = $True
    $button1.Text = "Run "
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 90
    $System_Drawing_Point.Y = 325
    $button1.Location = $System_Drawing_Point
    $button1.add_Click($button1Click)

    $form1.Controls.Add($button1)
    
    # Button 2
    
    $button2.Name = "button2"
    $button2.Width = 75
    $button2.Height = 25
    $button2.UseVisualStyleBackColor = $True
    $button2.Text = "Clear Host "
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 10
    $System_Drawing_Point.Y = 325
    $button2.Location = $System_Drawing_Point
    $button2.add_Click($button2Click)
 
    $form1.Controls.Add($button2)
 
    # List View (the script pane where results are displayed)

    $ListView.View = [System.Windows.Forms.View]::Details
    $ListView.Width = $form1.ClientRectangle.Width - 185
    $ListView.Height = $form1.ClientRectangle.Height - 15
    $ListView.Name = "listBox1"
    $ListView.Columns.Add("SCRIPT WINDOW", -2) | Out-Null
    $ListView.Font = "Lucida Console"
    $ListView.LabelEdit = $True
    $ListView.MultiSelect = $True
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 180
    $System_Drawing_Point.Y = 10
    $ListView.Location = $System_Drawing_Point

    $form1.Controls.Add($ListView)

    # Text Box (where you enter Computer Name)
    
    $textBox1.Width = 150
    $textBox1.Height = 25
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 10
    $System_Drawing_Point.Y = 27
    $textBox1.Location = $System_Drawing_Point
    $textBox1.Text

    $form1.Controls.Add($textBox1)

    # Label 1
    
    $label1.Text = "Computer Name:"
    $label1.Height = 15
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 10
    $System_Drawing_Point.Y = 10
    $label1.Location = $System_Drawing_Point

    $form1.Controls.Add($label1)

    # Label 2
    
    $label2.Text = "________________________________________________________________________________________________________________________"
    $label2.Height = 15
    $label2.Width = 190
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 0
    $System_Drawing_Point.Y = 130
    $label2.Location = $System_Drawing_Point

    $form1.Controls.Add($label2)
    
    # Check Box 7

    $checkBox7.UseVisualStyleBackColor = $True
    $checkBox7.Width = 105
    $checkBox7.Height = 25
    $checkBox7.Text = "Run a Script"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 10
    $System_Drawing_Point.Y = 190
    $checkBox7.Location = $System_Drawing_Point
    $checkBox7.DataBindings.DefaultDataSourceUpdateMode = 0
    $checkBox7.Name = "checkBox5"
 
    $form1.Controls.Add($checkBox7)

    # Check Box 6
    
    $checkBox6.UseVisualStyleBackColor = $True
    $checkBox6.Width = 150
    $checkBox6.Height = 25
    $checkBox6.Text = "Admin PowerShell ISE"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 10
    $System_Drawing_Point.Y = 170
    $checkBox6.Location = $System_Drawing_Point
    $checkBox6.DataBindings.DefaultDataSourceUpdateMode = 0
    $checkBox6.Name = "checkBox5"
 
    $form1.Controls.Add($checkBox6)

    # Check Box 5
    
    $checkBox5.UseVisualStyleBackColor = $True
    $checkBox5.Width = 150
    $checkBox5.Height = 25
    $checkBox5.Text = "Admin Command Prompt"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 10
    $System_Drawing_Point.Y = 150
    $checkBox5.Location = $System_Drawing_Point
    $checkBox5.DataBindings.DefaultDataSourceUpdateMode = 0
    $checkBox5.Name = "checkBox5"
 
    $form1.Controls.Add($checkBox5)

    # Check Box 4
    
    $checkBox4.UseVisualStyleBackColor = $True
    $checkBox4.Width = 150
    $checkBox4.Height = 25
    $checkBox4.Text = "Install Program"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 10
    $System_Drawing_Point.Y = 110
    $checkBox4.Location = $System_Drawing_Point
    $checkBox4.DataBindings.DefaultDataSourceUpdateMode = 0
    $checkBox4.Name = "checkBox4"
 
    $form1.Controls.Add($checkBox4)

    # Check Box 3
    
    $checkBox3.UseVisualStyleBackColor = $True
    $checkBox3.Width = 150
    $checkBox3.Height = 25
    $checkBox3.Text = "Enable WinRM"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 10
    $System_Drawing_Point.Y = 90
    $checkBox3.Location = $System_Drawing_Point
    $checkBox3.DataBindings.DefaultDataSourceUpdateMode = 0
    $checkBox3.Name = "checkBox3"
 
    $form1.Controls.Add($checkBox3)

    # Check Box 2
    
    $checkBox2.UseVisualStyleBackColor = $True
    $checkBox2.Width = 150
    $checkBox2.Height = 25
    $checkBox2.Text = "Logged on User"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 10
    $System_Drawing_Point.Y = 70
    $checkBox2.Location = $System_Drawing_Point
    $checkBox2.Name = "checkBox2"
 
    $form1.Controls.Add($checkBox2)
 
    # Check Box 1

    $checkBox1.UseVisualStyleBackColor = $True
    $checkBox1.Width = 150
    $checkBox1.Height = 25
    $checkBox1.Text = "System Info"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 10
    $System_Drawing_Point.Y = 50
    $checkBox1.Location = $System_Drawing_Point
    $checkBox1.DataBindings.DefaultDataSourceUpdateMode = 0
    $checkBox1.Name = "checkBox1"
 
    $form1.Controls.Add($checkBox1)
    # ----------------------------------------------------------------------------
    
    $InitialFormWindowState = $form1.WindowState
    $form1.add_Load($OnLoadForm_StateCorrection)

    $form1.ShowDialog() | Out-Null
}

GenerateForm

# OPtional Commands that have been removed

# Clears the host if included after button click action
#$ListView.Items.Clear(); 

# This displays an input poppup window that can be used as another method of getting input.
#$Computer = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter the Computer Name","Input Required",$env:COMPUTERNAME)