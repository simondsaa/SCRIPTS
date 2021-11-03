$LocalUser = Get-WmiObject Win32_ComputerSystem
$LocalEDI = $LocalUser.UserName.TrimStart("AREA52\")
$Code = $LocalEDI.Substring(6,4)

{
    Write-Host "Invalid code was entered... Closing"
    Start-Sleep -Seconds 2
    Exit
}

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
$checkBox19 = New-Object System.Windows.Forms.CheckBox
$checkBox18 = New-Object System.Windows.Forms.CheckBox
$checkBox17 = New-Object System.Windows.Forms.CheckBox
$checkBox16 = New-Object System.Windows.Forms.CheckBox
$checkBox15 = New-Object System.Windows.Forms.CheckBox
$checkBox14 = New-Object System.Windows.Forms.CheckBox
$checkBox13 = New-Object System.Windows.Forms.CheckBox
$checkBox12 = New-Object System.Windows.Forms.CheckBox
$checkBox11 = New-Object System.Windows.Forms.CheckBox
$checkBox10 = New-Object System.Windows.Forms.CheckBox
$checkBox9 = New-Object System.Windows.Forms.CheckBox
$checkBox8 = New-Object System.Windows.Forms.CheckBox
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
    $Places = $checkBox1, $checkBox2, $checkBox3, $checkBox4, $checkBox5, $checkBox6, $checkBox7, $checkBox8, $checkBox9, $checkBox10, $checkBox11, $checkBox12, $checkBox12, $checkBox14, $checkBox15, $checkBox16, $checkBox17, $checkBox19
    $Lunch = Get-Random -Count 1 $Places
    
    While ($Lunch.Checked)
    {
        $Lunch = Get-Random -Count 1 $Places
    }

    $c = New-Object -Comobject wscript.shell
    $b = $c.popup($Lunch.Text,0,"Lunch",0)

    $Body = "Lunchinator 9000 says: "+$Lunch.Text
    
    If ($checkBox17.Checked)
    {
        Send-MailMessage -From "lonnie.stringer.3@us.af.mil" -To "timothy.brady.11@us.af.mil", "lonnie.stringer.3@us.af.mil", "leej.tobler@us.af.mil", "ashley.thompson.21@us.af.mil", "john.fisher.26@us.af.mil", "christopher.chesnek@us.af.mil" -Priority High -Subject "Lunchinator 9000 Says:" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil
    }
}

# Generating the Form or Scripter Window
# ----------------------------------------------------------------------------

# Form 1

$form1.Text = "L 9000"
$form1.Name = "form1"
$form1.Width = 210
$form1.Height = 310
$form1.FormBorderStyle = "Fixed3D"
$form1.MaximizeBox = $false
$form1.MinimizeBox = $true

# Button 1

$button1.Name = "button1"
$button1.Width = 180
$button1.Height = 25
$button1.UseVisualStyleBackColor = $True
$button1.Text = "Run "
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 245
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

#$form1.Controls.Add($button2)

# Label 1

$label1.Text = "Check to Exclude Location"
$label1.Width = 150
$label1.Height = 15
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 10
$label1.Location = $System_Drawing_Point

$form1.Controls.Add($label1)

# Check Box 1

$checkBox1.UseVisualStyleBackColor = $True
$checkBox1.Width = 100
$checkBox1.Height = 25
$checkBox1.Text = "Bowling"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 30
$checkBox1.Location = $System_Drawing_Point
$checkBox1.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox1.Name = "checkBox1"

$form1.Controls.Add($checkBox1)

# Check Box 2

$checkBox2.UseVisualStyleBackColor = $True
$checkBox2.Width = 100
$checkBox2.Height = 25
$checkBox2.Text = "BX"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 50
$checkBox2.Location = $System_Drawing_Point
$checkBox2.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox2.Name = "checkBox2"

$form1.Controls.Add($checkBox2)

# Check Box 3

$checkBox3.UseVisualStyleBackColor = $True
$checkBox3.Width = 100
$checkBox3.Height = 25
$checkBox3.Text = "Chick-fil-a"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 70
$checkBox3.Location = $System_Drawing_Point
$checkBox3.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox3.Name = "checkBox3"

$form1.Controls.Add($checkBox3)

# Check Box 4

$checkBox4.UseVisualStyleBackColor = $True
$checkBox4.Width = 100
$checkBox4.Height = 25
$checkBox4.Text = "DQ"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 90
$checkBox4.Location = $System_Drawing_Point
$checkBox4.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox4.Name = "checkBox4"

$form1.Controls.Add($checkBox4)

# Check Box 5

$checkBox5.UseVisualStyleBackColor = $True
$checkBox5.Width = 100
$checkBox5.Height = 25
$checkBox5.Text = "Firehouse"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 110
$checkBox5.Location = $System_Drawing_Point
$checkBox5.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox5.Name = "checkBox5"

$form1.Controls.Add($checkBox5)

# Check Box 6

$checkBox6.UseVisualStyleBackColor = $True
$checkBox6.Width = 100
$checkBox6.Height = 25
$checkBox6.Text = "Gary's"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 130
$checkBox6.Location = $System_Drawing_Point
$checkBox6.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox6.Name = "checkBox6"

$form1.Controls.Add($checkBox6)

# Check Box 7

$checkBox7.UseVisualStyleBackColor = $True
$checkBox7.Width = 100
$checkBox7.Height = 25
$checkBox7.Text = "Hibachi"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 150
$checkBox7.Location = $System_Drawing_Point
$checkBox7.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox7.Name = "checkBox7"

$form1.Controls.Add($checkBox7)

# Check Box 8

$checkBox8.UseVisualStyleBackColor = $True
$checkBox8.Width = 100
$checkBox8.Height = 25
$checkBox8.Text = "Jimmy Johns"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 170
$checkBox8.Location = $System_Drawing_Point
$checkBox8.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox8.Name = "checkBox8"

$form1.Controls.Add($checkBox8)

# Check Box 9

$checkBox9.UseVisualStyleBackColor = $True
$checkBox9.Width = 100
$checkBox9.Height = 25
$checkBox9.Text = "JR's"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 30
$checkBox9.Location = $System_Drawing_Point
$checkBox9.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox9.Name = "checkBox9"

$form1.Controls.Add($checkBox9)

# Check Box 10

$checkBox10.UseVisualStyleBackColor = $True
$checkBox10.Width = 100
$checkBox10.Height = 25
$checkBox10.Text = "McDonalds"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 50
$checkBox10.Location = $System_Drawing_Point
$checkBox10.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox10.Name = "checkBox10"

$form1.Controls.Add($checkBox10)

# Check Box 11

$checkBox11.UseVisualStyleBackColor = $True
$checkBox11.Width = 100
$checkBox11.Height = 25
$checkBox11.Text = "Napoli's"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 70
$checkBox11.Location = $System_Drawing_Point
$checkBox11.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox11.Name = "checkBox11"

$form1.Controls.Add($checkBox11)

# Check Box 12

$checkBox12.UseVisualStyleBackColor = $True
$checkBox12.Width = 100
$checkBox12.Height = 25
$checkBox12.Text = "Old Mexico"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 90
$checkBox12.Location = $System_Drawing_Point
$checkBox12.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox12.Name = "checkBox12"

$form1.Controls.Add($checkBox12)

# Check Box 13

$checkBox13.UseVisualStyleBackColor = $True
$checkBox13.Width = 100
$checkBox13.Height = 25
$checkBox13.Text = "Sonic"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 110
$checkBox13.Location = $System_Drawing_Point
$checkBox13.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox13.Name = "checkBox13"

$form1.Controls.Add($checkBox13)

# Check Box 14

$checkBox14.UseVisualStyleBackColor = $True
$checkBox14.Width = 100
$checkBox14.Height = 25
$checkBox14.Text = "Thai"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 130
$checkBox14.Location = $System_Drawing_Point
$checkBox14.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox14.Name = "checkBox14"

$form1.Controls.Add($checkBox14)

# Check Box 15

$checkBox15.UseVisualStyleBackColor = $True
$checkBox15.Width = 100
$checkBox15.Height = 25
$checkBox15.Text = "Third Party"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 150
$checkBox15.Location = $System_Drawing_Point
$checkBox15.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox15.Name = "checkBox15"

$form1.Controls.Add($checkBox15)

# Check Box 16

$checkBox16.UseVisualStyleBackColor = $True
$checkBox16.Width = 100
$checkBox16.Height = 25
$checkBox16.Text = "Wendy's"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 170
$checkBox16.Location = $System_Drawing_Point
$checkBox16.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox16.Name = "checkBox16"

$form1.Controls.Add($checkBox16)

# Check Box 17

$checkBox17.UseVisualStyleBackColor = $True
$checkBox17.Width = 100
$checkBox17.Height = 25
$checkBox17.Text = "Subway"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 190
$checkBox17.Location = $System_Drawing_Point
$checkBox17.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox17.Name = "checkBox17"

$form1.Controls.Add($checkBox17)

# Check Box 18

$checkBox18.UseVisualStyleBackColor = $True
$checkBox18.Width = 100
$checkBox18.Height = 25
$checkBox18.Text = "Send Email"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 60
$System_Drawing_Point.Y = 220
$checkBox18.Location = $System_Drawing_Point
$checkBox18.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox18.Name = "checkBox18"

$form1.Controls.Add($checkBox18)

# Check Box 19

$checkBox19.UseVisualStyleBackColor = $True
$checkBox19.Width = 100
$checkBox19.Height = 25
$checkBox19.Text = "Oasis"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 110
$System_Drawing_Point.Y = 190
$checkBox19.Location = $System_Drawing_Point
$checkBox19.DataBindings.DefaultDataSourceUpdateMode = 0
$checkBox19.Name = "checkBox19"

$form1.Controls.Add($checkBox19)

# ----------------------------------------------------------------------------

$InitialFormWindowState = $form1.WindowState
$form1.add_Load($OnLoadForm_StateCorrection)

$form1.ShowDialog() | Out-Null