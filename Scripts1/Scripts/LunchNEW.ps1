[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
 [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

 #This creates the form and sets its size and position
 $objForm = New-Object System.Windows.Forms.Form 
 $objForm.Text = "Bowling Grub"
 $objForm.Size = New-Object System.Drawing.Size(300,615) 
 $objForm.StartPosition = "CenterScreen"

 $objForm.KeyPreview = $True
 $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
     {$empID=$objTextBox1.Text;$sn=$objTextBox2.Text;$gn=$objTextBox3.Text;$email=$objTextBox4.Text;$title=$objDepartmentListbox.SelectedItem;
      $office=$objOfficeListbox.SelectedItem;$objForm.Close()}})
 $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
     {$objForm.Close()}})

 #This creates a label for the TextBox1
 $objLabel1 = New-Object System.Windows.Forms.Label
 $objLabel1.Location = New-Object System.Drawing.Size(10,20) 
 $objLabel1.Size = New-Object System.Drawing.Size(280,20) 
 $objLabel1.Text = "Name:"
 $objForm.Controls.Add($objLabel1) 
  
 #This creates the TextBox1
 $objTextBox1 = New-Object System.Windows.Forms.TextBox 
 $objTextBox1.Location = New-Object System.Drawing.Size(10,40) 
 $objTextBox1.Size = New-Object System.Drawing.Size(260,20)
 $objTextBox1.TabIndex = 0 
 $objForm.Controls.Add($objTextBox1)

 #This creates a label for the TextBox2
 $objLabel1 = New-Object System.Windows.Forms.Label
 $objLabel1.Location = New-Object System.Drawing.Size(10,20) 
 $objLabel1.Size = New-Object System.Drawing.Size(280,20) 
 $objLabel1.Text = "What would you like to eat?"
 $objForm.Controls.Add($objLabel2) 

 #This creates the TextBox2
 $objTextBox1 = New-Object System.Windows.Forms.TextBox 
 $objTextBox1.Location = New-Object System.Drawing.Size(10,40) 
 $objTextBox1.Size = New-Object System.Drawing.Size(260,20)
 $objTextBox1.TabIndex = 1 
 $objForm.Controls.Add($objTextBox2)

 #This creates a label for the TextBox3
 $objLabel2 = New-Object System.Windows.Forms.Label
 $objLabel2.Location = New-Object System.Drawing.Size(10,70) 
 $objLabel2.Size = New-Object System.Drawing.Size(280,20) 
 $objLabel2.Text = "With everything? Minus something?"
 $objForm.Controls.Add($objLabel3)  

 #This creates the TextBox3
 $objTextBox2 = New-Object System.Windows.Forms.TextBox 
 $objTextBox2.Location = New-Object System.Drawing.Size(10,90) 
 $objTextBox2.Size = New-Object System.Drawing.Size(260,20)
 $objTextBox2.TabIndex = 2  
 $objForm.Controls.Add($objTextBox3)

 #This creates a label for the TextBox4
 $objLabel3 = New-Object System.Windows.Forms.Label
 $objLabel3.Location = New-Object System.Drawing.Size(10,120) 
 $objLabel3.Size = New-Object System.Drawing.Size(280,20) 
 $objLabel3.Text = "Do you want a drink?"
 $objForm.Controls.Add($objLabel4)  

 #This creates the TextBox4
 $objTextBox3 = New-Object System.Windows.Forms.TextBox 
 $objTextBox3.Location = New-Object System.Drawing.Size(10,140) 
 $objTextBox3.Size = New-Object System.Drawing.Size(260,20)
 $objTextBox3.TabIndex = 3
 $objForm.Controls.Add($objTextBox4)

 #This creates a label for the TextBox5
 $objLabel4 = New-Object System.Windows.Forms.Label
 $objLabel4.Location = New-Object System.Drawing.Size(10,170) 
 $objLabel4.Size = New-Object System.Drawing.Size(280,20) 
 $objLabel4.Text = "Winner Prediction:"
 $objForm.Controls.Add($objLabel5)  

 #This creates the TextBox5
 $objTextBox4 = New-Object System.Windows.Forms.TextBox 
 $objTextBox4.Location = New-Object System.Drawing.Size(10,190) 
 $objTextBox4.Size = New-Object System.Drawing.Size(260,20)
 $objTextBox4.TabIndex = 4
 $objForm.Controls.Add($objTextBox5)

  #This creates a label for the TextBox6
 $objLabel4 = New-Object System.Windows.Forms.Label
 $objLabel4.Location = New-Object System.Drawing.Size(10,170) 
 $objLabel4.Size = New-Object System.Drawing.Size(280,20) 
 $objLabel4.Text = "La-Who-Za-Her Prediction:"
 $objForm.Controls.Add($objLabel6)  

 #This creates the TextBox6
 $objTextBox4 = New-Object System.Windows.Forms.TextBox 
 $objTextBox4.Location = New-Object System.Drawing.Size(10,190) 
 $objTextBox4.Size = New-Object System.Drawing.Size(260,20)
 $objTextBox4.TabIndex = 5
 $objForm.Controls.Add($objTextBox6)

 #This creates a checkbox called Employee
 $objTypeCheckbox = New-Object System.Windows.Forms.Checkbox 
 $objTypeCheckbox.Location = New-Object System.Drawing.Size(10,220) 
 $objTypeCheckbox.Size = New-Object System.Drawing.Size(500,20)
 $objTypeCheckbox.Text = "I agree to not talk about work"
 $objTypeCheckbox.TabIndex = 6
 $objForm.Controls.Add($objTypeCheckbox)

 #This creates a checkbox called Citrix User
 $objCitrixUserCheckbox = New-Object System.Windows.Forms.Checkbox 
 $objCitrixUserCheckbox.Location = New-Object System.Drawing.Size(10,240) 
 $objCitrixUserCheckbox.Size = New-Object System.Drawing.Size(500,20)
 $objCitrixUserCheckbox.Text = "I agree to not be a Debby Downer"
 $objCitrixUserCheckbox.TabIndex = 7
 $objForm.Controls.Add($objCitrixUserCheckbox)

 #This creates a checkbox called Non-Citrix User
 $objNonCitrixUserCheckbox = New-Object System.Windows.Forms.Checkbox 
 $objNonCitrixUserCheckbox.Location = New-Object System.Drawing.Size(10,260) 
 $objNonCitrixUserCheckbox.Size = New-Object System.Drawing.Size(500,20)
 $objNonCitrixUserCheckbox.Text = "I agree to make fart noises when Jon is throwing"
 $objNonCitrixUserCheckbox.TabIndex = 8
 $objForm.Controls.Add($objNonCitrixUserCheckbox)

 #This creates the Ok button and sets the event
 $OKButton = New-Object System.Windows.Forms.Button
 $OKButton.Location = New-Object System.Drawing.Size(120,540)
 $OKButton.Size = New-Object System.Drawing.Size(75,23)
 $OKButton.Text = "OK"
 $OKButton.Add_Click({$empID=$objTextBox1.Text;$sn=$objTextBox2.Text;$gn=$objTextBox3.Text;$email=$objTextBox4.Text;$title=$objDepartmentListbox.SelectedItem;
                      $office=$objOfficeListbox.SelectedItem;$objForm.Close()})
 $OKButton.TabIndex = 9
 $objForm.Controls.Add($OKButton)

 #This creates the Cancel button and sets the event
 $CancelButton = New-Object System.Windows.Forms.Button
 $CancelButton.Location = New-Object System.Drawing.Size(195,540)
 $CancelButton.Size = New-Object System.Drawing.Size(75,23)
 $CancelButton.Text = "Cancel"
 $CancelButton.Add_Click({$objForm.Close()})
 $CancelButton.TabIndex = 10
 $objForm.Controls.Add($CancelButton)

 $objForm.Add_Shown({$objForm.Activate()})
 [void] $objForm.ShowDialog()


 #Combine last name with first name to create the Display Name
 $dn = "$sn, $gn"