# Written by SSgt Timothy Brady
# Tyndall AFB, Panama City, FL
# Created February 5, 2016

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
$button3 = New-Object System.Windows.Forms.Button
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
$button2Click=
{
    $ListView.Items.Add($button2.Text).BackColor = "Silver"
    $Message = $null    
    $Letters = $textBox1.Text

    $Enc = [System.Text.Encoding]::UTF8
    $Letters = $Enc.GetBytes($Letters)
    
    ForEach ($Letter in $Letters)
    {
        If ($Letter -eq "32")
        {
            $Letter = $Letter-1
        }
    
        $NewLetter = $Letter+1
        $New = [char]$NewLetter
        $Message += $New
    }

    $ListView.Items.Add("$Message")    
}
     
$button3Click=
{
    $ListView.Items.Add($button3.Text).BackColor = "Silver"
    $DMessage = $null
    $DLetters = $textBox1.Text

    $Enc = [System.Text.Encoding]::UTF8
    $DLetters = $Enc.GetBytes($DLetters)
    
    ForEach ($DLetter in $DLetters)
    {
        If ($DLetter -eq "32")
        {
            $DLetter = $DLetter+1
        }
            $DNewLetter = $DLetter-1
            $DNew = [char]$DNewLetter
            $DMessage += $DNew
    }

    $ListView.Items.Add("$DMessage")
}


# "Clear Host" Button Action
# ----------------------------------------------------------------------------
$button1Click=
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
    
$form1.Text = "Crypto Script"
$form1.Name = "form1"
$form1.Width = 500
$form1.Height = 200
$form1.FormBorderStyle = "Fixed3D"
$form1.MaximizeBox = $false
$form1.MinimizeBox = $true

# Button 1
    
$button1.Name = "button1"
$button1.Width = 150
$button1.Height = 25
$button1.UseVisualStyleBackColor = $True
$button1.Text = "Clear Host"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 125
$button1.Location = $System_Drawing_Point
$button1.add_Click($button1Click)

$form1.Controls.Add($button1)
    
# Button 2
    
$button2.Name = "button2"
$button2.Width = 75
$button2.Height = 25
$button2.UseVisualStyleBackColor = $True
$button2.Text = "Encrypt"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 50
$button2.Location = $System_Drawing_Point
$button2.add_Click($button2Click)
 
$form1.Controls.Add($button2)

# Button 3
    
$button3.Name = "button3"
$button3.Width = 75
$button3.Height = 25
$button3.UseVisualStyleBackColor = $True
$button3.Text = "Decrypt"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 90
$System_Drawing_Point.Y = 50
$button3.Location = $System_Drawing_Point
$button3.add_Click($button3Click)
 
$form1.Controls.Add($button3)
 
# List View (the script pane where results are displayed)

$ListView.View = [System.Windows.Forms.View]::Details
$ListView.Width = $form1.ClientRectangle.Width - 185
$ListView.Height = $form1.ClientRectangle.Height - 15
$ListView.Name = "listBox1"
$ListView.Columns.Add("Results Window", -2) | Out-Null
$ListView.Font = "Lucida Console"
$ListView.LabelEdit = $True
$ListView.MultiSelect = $True
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 180
$System_Drawing_Point.Y = 10
$ListView.Location = $System_Drawing_Point

$form1.Controls.Add($ListView)

# Text Box
    
$textBox1.Width = 155
$textBox1.Height = 100
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 27
$textBox1.Location = $System_Drawing_Point
$textBox1.Text

$form1.Controls.Add($textBox1)

# Label 1
    
$label1.Text = "Input Message"
$label1.Height = 15
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 10
$label1.Location = $System_Drawing_Point

$form1.Controls.Add($label1)
    
# Check Box 2
    
$checkBox2.UseVisualStyleBackColor = $True
$checkBox2.Width = 150
$checkBox2.Height = 25
$checkBox2.Text = "Decrypt"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 70
$checkBox2.Location = $System_Drawing_Point
$checkBox2.Name = "checkBox2"
 
#$form1.Controls.Add($checkBox2)
 
# Check Box 1

$checkBox1.UseVisualStyleBackColor = $True
$checkBox1.Width = 150
$checkBox1.Height = 25
$checkBox1.Text = "Encrypt"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 50
$checkBox1.Location = $System_Drawing_Point
$checkBox1.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox1.Name = "checkBox1"
 
#$form1.Controls.Add($checkBox1)
# ----------------------------------------------------------------------------
    
$InitialFormWindowState = $form1.WindowState
$form1.add_Load($OnLoadForm_StateCorrection)

$form1.ShowDialog() | Out-Null

# OPtional Commands that have been removed

# Clears the host if included after button click action
#$ListView.Items.Clear(); 

# This displays an input poppup window that can be used as another method of getting input.
#$Computer = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter the Computer Name","Input Required",$env:COMPUTERNAME)