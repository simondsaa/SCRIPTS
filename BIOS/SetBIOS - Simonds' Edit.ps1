
<# 
 .Synopsis 
     The Set-BIOS function allows you to change BIOS settings from your local or a remote computer or also multiple remote computers. 
     You can change BIOS for the following manufacturer: Dell, HP, Lenovo. 
     Settings to change should be located in a CSV file. 
     If you need other manufacturer send me a mail at: damien.vanrobaeys@gmail.com 
      
 .DESCRIPTION 
     The Set-BIOS function allows you to change BIOS settings from your local or a remote computer or also multiple remote computers. 
     You can change BIOS for the following manufacturer: Dell, HP, Lenovo 
      
     Settings to change should be located in a CSV file. 
     The CSV file should contain two columns: Setting and Value, see below an example: 
     Setting,Value 
     Fast Boot,Enable 
     Select Language, Francais 
     Audio Alerts During Boot, Disable 
      
     See below available parameters: 
     -CSV "CSVPath": Path of the CSV file containing BIOS settings to change 
     -Computer "ComputerName": Remote computer name on which you want to change settings 
     -Multiple: Use it to change settings on multiple remote computers (Do not use -Computer with it) 
     -Manufacturer "HP / Dell, Lenovo": To use with the -Multiple parameter. Type the manufacturer of the list of remote computers: Dell, HP or Lenovo. 
     -List "ComputersList.txt": To use with the -Multiple parameter. Path of the TXT file that contains the list of remote computer name 
     -Password: If a BIOS is setted. You will have to type the password once it's asked 
  
 .EXAMPLE 
 PS Root\> Set-BIOS 
 The command above will change BIOS settings on your local computer. 
 It will ask for the CSV path as below: 
 CSV PATH - Please enter the CSV path containing settings to change with new values: 
  
 .EXAMPLE 
 PS Root\> Set-BIOS -CSV "C:\BIOS_Settings.csv" 
 The command above will change BIOS settings on your local computer and apply settings values located in the file C:\BIOS_Settings.csv 
  
 .EXAMPLE 
 PS Root\> Set-BIOS Computer "MyComputer" -CSV "C:\BIOS_Settings.csv" 
 The command above will change BIOS settings on the computer MyComputer and apply settings values located in the file C:\BIOS_Settings.csv 
 It will prompt you to enter credentials to connect to the remote computer 
  
 .EXAMPLE 
 PS Root\> Set-BIOS -Computer "MyComputer" -CSV "C:\BIOS_Settings.csv" -Password PASSWORD 
 The command above will change BIOS settings on the computer MyComputer and apply settings values located in the file C:\BIOS_Settings.csv 
 It will ask for the BIOS password as below: 
 PASSWORD - Please enter your BIOS password: ************ 
  
 .EXAMPLE 
 PS Root\> Set-BIOS -multiple -Manufacturer HP -List "C:\ComputersList.txt" -CSV "C:\BIOS_Settings.csv" 
 The command above will change BIOS settings for the list of HP computers located in C:\ComputersList.txt 
 It will apply setting values located in the file C:\BIOS_Settings.csv 
  
 .NOTES 
     Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com 
 #>

function Set-Bios
{
    [CmdletBinding()]
    Param(
            [string]$Computer,
            [switch]$Password,    
            [switch]$Multiple,    
            [string]$List,
            [string]$Manufacturer,            
            [string]$CSV            
         )

    Begin
    {
        # Check if user has selected multiple switch
        If(($Multiple) -and ($Computer -ne ""))
            {
                write-host ""
                write-host "###############################################################" -Foreground yellow
                write-host "The Computer parameter can't be used with the muliple parameter" -Foreground yellow
                write-host "###############################################################" -Foreground yellow    
                write-host ""
                break                    
            }    
            
        # If user has selected both List and Computer parameter
        If(($List -ne "") -and ($Computer -ne ""))
            {
                write-host ""
                write-host "###############################################################" -Foreground yellow
                write-host "The Computer parameter can't be used with the muliple parameter" -Foreground yellow
                write-host "###############################################################" -Foreground yellow    
                write-host ""
                break                    
            }                
    
        # If user has typed a computer name
        If(($Computer -ne ""))
            {    
                # Check connection to the remote connection
                $Check_Connection = test-connection -computername $Computer -quiet -count 1
                If($Check_Connection -eq $true)
                    {
                       
                        $Get_Remote_Vendor = {(Get-WmiObject Win32_Computersystem).manufacturer}
                        $Script:Get_manufacturer = Invoke-Command -ComputerName $Computer -ScriptBlock $Get_Remote_Vendor -ErrorVariable errmsg 2>$null    
                        If($errmsg -ne $null)
                            {
                                write-host "KO ==> Connexion KO on $Computer (Check your credentials)" -Foreground yellow          
                                break                                            
                            }
                    }
                Else
                    {
                        write-host "KO ==> Connexion KO on $Computer" -Foreground yellow              
                        break
                    }    
            }
        Else
            {
                # In this case user hasn't typed a computer name, meaning we will set bios on local computer
                $Script:Get_manufacturer = (Get-WmiObject Win32_Computersystem).manufacturer
            }
            
        # if user has selected password switch, we will use a BIOS password
        If($Password)
            {    
                $Script:BIOS_PWD = Read-Host -assecurestring "PASSWORD - Please enter your BIOS password"
                $Script:MyPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($BIOS_PWD))
            }    

        # If user has typed a CSV path
        If(($CSV -ne ""))
            {    
                $Script:Settings_CSV = $CSV
            }
        Else
            {
                # If not we will ask for the CSV path
                $Script:Settings_CSV = $CSV
                $Script:Settings_CSV = read-host "CSV PATH - Please enter the CSV path containing settings to change with new values"
            }
            
        # Test if CSV path exists
        If(!(test-path $Settings_CSV))
            {
                write-host "KO ==> Can not find file $Settings_CSV" -Foreground yellow              
                break                    
            }            

        # if user has selected the Multiple switch
        If($Multiple)
            {        
                If($Manufacturer -ne "")
                    {
                        $Script:Get_manufacturer = $Manufacturer
                    }
                Else
                    {
                        $Script:Get_manufacturer = read-host "MANUFACTURER - Please enter the manufacturer of remote computer"                     
                    }
                    
                If($List -ne "") 
                    {
                        $Script:Computers_List = $List
                    }
                Else
                    {
                        $Script:Computers_List = read-host "COMPUTERS LIST - Please enter the path of the TXT file containing computers list" 
                    }
                    
                If(!(test-path $Settings_CSV))
                    {
                        write-host "KO ==> Can not find file $Computers_List" -Foreground yellow              
                        break                    
                    }                        
                    
                
            }                
            

        Function Set_Dell_BIOS_Settings
            {                        
                param
                (
                    [array]$CSV_Parameters,
                    [string]$MyPassword    
                )                
                
                If(($Muliple) -or ($Computer -ne ""))
                    {
                        $Exported_CSV = "c:\temp\Exported_BIOS_Settings.csv"                                                                        
                        $CSV_Parameters | out-file -LiteralPath $Exported_CSV
                        $Get_CSV_Content = import-csv $Exported_CSV                                
                    }
                Else
                    {
                        $Script:Get_CSV_Content = import-csv $Settings_CSV                    
                    }

                $WarningPreference='silentlycontinue'
                If (Get-Module -ListAvailable -Name DellBIOSProvider)
                    {} 
                Else 
                    {
                        Install-Module -Name DellBIOSProvider -Force 
                    }         
                get-command -module DellBIOSProvider | out-null            

                $IsPasswordSet = (Get-Item -Path DellSmbios:\Security\IsAdminPasswordSet).currentvalue    
                If ((($MyPassword -eq $null) -or ($MyPassword -eq "")) -and ($IsPasswordSet -eq $true))                                
                    {
                        write-host ""
                        write-host "###############################################################" -Foreground yellow
                        write-host " A password is configured in your BIOS ($env:computername)" -Foreground yellow
                        write-host " Please add the -password parameter" -Foreground yellow                
                        write-host "###############################################################" -Foreground yellow    
                        write-host ""
                        break                                        
                    }
                    
                $Dell_BIOS = get-childitem -path DellSmbios:\ | foreach {
                get-childitem -path @("DellSmbios:\" + $_.Category)  | select-object attribute, currentvalue, possiblevalues, PSChildName
                }                         
                    
                ForEach($New_Setting in $Get_CSV_Content)
                    {    
                        $Setting_To_Set = $New_Setting.Setting    
                        $Setting_NewValue_To_Set = $New_Setting.Value    
                        ForEach($Current_Setting in $Dell_BIOS | Where {$_.attribute -eq $Setting_To_Set})
                            {    
                                $Attribute = $Current_Setting.attribute
                                $Setting_Cat = $Current_Setting.PSChildName
                                $Setting_Current_Value = $Current_Setting.CurrentValue
                
                                If (($IsPasswordSet -eq $true))
                                    {            
                                        $Password_To_Use = $MyPassword
                                        # & Set-Item -Path Dellsmbios:\$Setting_Cat\$Attribute -Value $Setting_NewValue_To_Set -Password $Password_To_Use
                                        & Set-Item -Path Dellsmbios:\$Setting_Cat\$Attribute -Value $Setting_NewValue_To_Set -Password $Password_To_Use -errorvariable ohgodanerror -ea silentlycontinue    
                                        If($ohgodanerror -ne $null)
                                            {
                                                write-host "KO ==> Can not change setting $Attribute (See below)" -Foreground Yellow    
                                                $ohgodanerror 
                                            }
                                        Else
                                            {
                                                write-host "OK ==> New value for $Attribute is $Setting_NewValue_To_Set"                                            
                                            }                                        
                                    }
                                Else
                                    {
                                        # & Set-Item -Path Dellsmbios:\$Setting_Cat\$Attribute -Value $Setting_NewValue_To_Set
                                        & Set-Item -Path Dellsmbios:\$Setting_Cat\$Attribute -Value $Setting_NewValue_To_Set -errorvariable ohgodanerror -ea silentlycontinue    
                                        If($ohgodanerror -ne $null)
                                            {
                                                write-host "KO ==> Can not change setting $Attribute (See below)" -Foreground Yellow    
                                                $ohgodanerror 
                                            }
                                        Else
                                            {
                                                write-host "OK ==> New value for $Attribute is $Setting_NewValue_To_Set"                                            
                                            }
                                    }                            
                            }                            
                    }    
                    
                If(($Muliple) -or ($Computer -ne ""))
                    {
                        remove-item $Exported_CSV -force                
                    }            
            }        


        Function Set_HP_BIOS_Settings
            {    
                Param
                (
                    [array]$CSV_Parameters,
                    [string]$MyPassword    
                )                
                
                If(($Muliple) -or ($Computer -ne ""))
                    {
                        #$Exported_CSV = "c:\temp\Exported_BIOS_Settings.csv"                                                                        
                        #$CSV_Parameters | out-file -LiteralPath $Exported_CSV
                        #$Get_CSV_Content = import-csv $Exported_CSV                        
                    }
                Else
                    {
                        $Script:Get_CSV_Content = import-csv $Settings_CSV                                                
                    }
                
                $IsPasswordSet = (Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where {$_.Name -eq "Setup Password"}).IsSet                            
                If ((($MyPassword -eq $null) -or ($MyPassword -eq "")) -and ($IsPasswordSet -eq $true))                
                    {
                        write-host ""
                        write-host "###############################################################" -Foreground yellow
                        write-host " A password is configured in your BIOS ($env:computername)" -Foreground yellow
                        write-host " Please add the -password parameter" -Foreground yellow                
                        write-host "###############################################################" -Foreground yellow    
                        write-host ""
                        break                                        
                    }                

                $bios = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface
                ForEach($Settings in $Get_CSV_Content)
                    {
                        $MySetting = $Settings.Setting
                        $NewValue = $Settings.Value        

                        If (($IsPasswordSet -eq 1))
                            {            
                                $Password_To_Use = "<utf-16/>"+$MyPassword
                                $Execute_Change_Action = $bios.setbiossetting("$MySetting", "$NewValue",$Password_To_Use)     
                                $Change_Return_Code = $Execute_Change_Action.return
                                If(($Change_Return_Code) -eq 0)                                
                                    {
                                        write-host "OK ==> New value for $MySetting is $NewValue"
                                    }
                                Else
                                    {
                                        write-host "KO ==> Can not change setting $MySetting (Return code $Change_Return_Code)" -Foreground Yellow
                                    }
                            }
                        Else
                            {
                                $Execute_Change_Action = $bios.setbiossetting("$MySetting", "$NewValue","")
                                $Change_Return_Code = $Execute_Change_Action.return
                                If(($Change_Return_Code) -eq 0)                                
                                    {
                                        write-host "OK ==> New value for $MySetting is $NewValue"
                                    }
                                Else
                                    {
                                        write-host "KO ==> Can not change setting $MySetting (Return code $Change_Return_Code)" -Foreground Yellow
                                    }                                
                            }
                    }    
                    
                If(($Muliple) -or ($Computer -ne ""))
                    {
                        #remove-item $Exported_CSV -force                
                    }
            }


    
        Function Set_Lenovo_BIOS_Settings
            {
                Param
                (
                    [array]$CSV_Parameters,
                    [string]$MyPassword                                        
                )        

                If(($Muliple) -or ($Computer -ne ""))
                    {
                        $Exported_CSV = "c:\temp\Exported_BIOS_Settings.csv"                                                                        
                        $CSV_Parameters | out-file -LiteralPath $Exported_CSV
                        $Get_CSV_Content = import-csv $Exported_CSV                        
                    }
                Else
                    {
                        $Script:Get_CSV_Content = import-csv $Settings_CSV                    
                    }

                $IsPasswordSet = (gwmi -Class Lenovo_BiosPasswordSettings -Namespace root\wmi).PasswordState
                If ((($MyPassword -eq $null) -or ($MyPassword -eq "")) -and ($IsPasswordSet -eq 2))                                
                    {
                        write-host ""
                        write-host "###############################################################" -Foreground yellow
                        write-host " A password is configured in your BIOS ($env:computername)" -Foreground yellow
                        write-host " Please add the -password parameter" -Foreground yellow                
                        write-host "###############################################################" -Foreground yellow    
                        write-host ""
                        break                                        
                    }    

                $BIOS = gwmi -class Lenovo_SetBiosSetting -namespace root\wmi
                $SAVE_BIOS = (gwmi -class Lenovo_SaveBiosSettings -namespace root\wmi)
                ForEach($Settings in $Get_CSV_Content)
                    {
                        $MySetting = $Settings.Setting
                        $NewValue = $Settings.Value        
                        
                        If (($IsPasswordSet -eq 2))
                            {                                        
                                $Script:Password_To_Use = $MyPassword                                    
                                $Execute_Change_Action = $bios.SetBiosSetting("$MySetting,$NewValue,$Password_To_Use,ascii,us") 
                                $Change_Return_Code = $Execute_Change_Action.return 
                                If(($Change_Return_Code) -eq "Success")                                
                                    {
                                        write-host "OK ==> New value for $MySetting is $NewValue"
                                    }
                                Else
                                    {
                                        write-host "KO ==> Can not change setting $MySetting (Return code $Change_Return_Code)" -Foreground Yellow
                                    }                                
                            }
                        Else
                            {
                                $Execute_Change_Action = $BIOS.SetBiosSetting("$MySetting,$NewValue") 
                                $Change_Return_Code = $Execute_Change_Action.return 
                                If(($Change_Return_Code) -eq "Success")                                
                                    {
                                        write-host "OK ==> New value for $MySetting is $NewValue"
                                    }
                                Else
                                    {
                                        write-host "KO ==> Can not change setting $MySetting (Return code $Change_Return_Code)" -Foreground Yellow
                                    }                                    
                            }                        
                    }
                    
                If (($IsPasswordSet -eq 2))
                    {    
                        $SAVE_BIOS.SaveBiosSettings("$Password_To_Use,ascii,us")                                                                        
                    }
                Else
                    {
                        $SAVE_BIOS.SaveBiosSettings()                                                    
                    }
                    
                If(($Muliple) -or ($Computer -ne ""))
                    {
                        remove-item $Exported_CSV -force                
                    }                    
            }

        If($Get_manufacturer -like "*dell*")
            {
                $manufacturer = "Dell"
            }
        ElseIf($Get_manufacturer -like "*lenovo*")
            {
                $manufacturer = "Lenovo"
            }
        ElseIf(($Get_manufacturer -like "*HP*") -or ($Get_manufacturer -like "*hewlet*"))
            {
                $manufacturer = "HP"
            }
        ElseIf($Get_manufacturer -like "*toshiba*")
            {
                # $manufacturer = "Toshiba"
                write-host ""
                write-host "########################################################" -Foreground yellow
                write-host " Your manufacturer will be soon supported" -Foreground yellow
                write-host "########################################################" -Foreground yellow    
                write-host ""
                break                
            }
        Else
            {
                write-host ""
                write-host "########################################################" -Foreground yellow
                write-host " Your manufacturer is not supported by the module" -Foreground yellow
                write-host " Supported manufacturer: Dell, HP, Lenovo, Toshiba" -Foreground yellow                
                write-host "########################################################" -Foreground yellow    
                write-host ""
                break
            }
    }

    Process
    {
        write-host ""
        write-host "###################################################" -Foreground Cyan
        write-host " Your manufacturer is $manufacturer" -Foreground Cyan
        write-host "###################################################" -Foreground Cyan    
        write-host ""        

       switch ($manufacturer)
       {
           'Dell'
           {
                $Global:CSV_Content = get-content $Settings_CSV    

                If($Multiple) # Meaning you want to export from a list of computer
                    {
                        $Get_List_Content = Get-Content $Computers_List
                        $Script:Dest_Path = "c:\windows\temp"                            
                        ForEach($Computer_Line in $Get_List_Content)
                            {
                            
                                write-host ""
                                write-host "Working on $Computer_Line" -Foreground green     
                                $Check_Connection = test-connection -computername $Computer_Line -quiet -count 1
                                If($Check_Connection -eq $true)
                                    {
                                        Invoke-Command -ComputerName $Computer_Line -ScriptBlock ${Function:Set_Dell_BIOS_Settings} -ArgumentList (,$CSV_Content), $MyPassword                                                                         
                                    }
                                Else
                                    {
                                        write-host "Connexion KO on $Computer_Line" -Foreground yellow                                  
                                    }                                
                            
                            }
                    }
                ElseIf(($Computer -ne ""))
                    {                        
                        Invoke-Command -ComputerName $Computer -ScriptBlock ${Function:Set_Dell_BIOS_Settings} -ArgumentList (,$CSV_Content), $MyPassword                                 
                    }
                Else
                    {
                        Set_Dell_BIOS_Settings $CSV_Content $MyPassword    
                    }                    

           }
           'HP'
           {
             Try
                {        
                    $Script:CSV_Content = get-content $Settings_CSV    
                    If($Multiple) # Meaning you want to export from a list of computer
                        {                                        
                            $Get_List_Content = Get-Content $Computers_List
                            $Script:Dest_Path = "c:\windows\temp"    
                            ForEach($Computer_Line in $Get_List_Content)
                                {                                
                                    write-host ""
                                    write-host "Working on $Computer_Line" -Foreground green     
                                    $Check_Connection = test-connection -computername $Computer_Line -quiet -count 1
                                    If($Check_Connection -eq $true)
                                        {
                                            Invoke-Command -ComputerName $Computer_Line -ScriptBlock ${Function:Set_HP_BIOS_Settings} -ArgumentList (,$CSV_Content), $MyPassword        
                                        }
                                    Else
                                        {
                                            write-host "Connexion KO on $Computer_Line" -Foreground yellow                                  
                                        }                                    
                                }
                        }                            
                    ElseIf(($Computer -ne ""))
                        {                        

                            Invoke-Command -ComputerName $Computer -ScriptBlock ${Function:Set_HP_BIOS_Settings} -ArgumentList (,$CSV_Content), $MyPassword                                      
                        }
                    Else
                        {                        
                            Set_HP_BIOS_Settings $CSV_Content $MyPassword    
                        }
                }
                    Catch
                    {}
           }
           'Lenovo'
           {
                $Global:CSV_Content = get-content $Settings_CSV    
                If($Multiple) # Meaning you want to export from a list of computer
                    {
                        $Get_List_Content = Get-Content $Computers_List
                        $Script:Dest_Path = "c:\windows\temp"                            
                        ForEach($Computer_Line in $Get_List_Content)
                            {
                                write-host ""
                                write-host "Working on $Computer_Line" -Foreground green                
                                $Check_Connection = test-connection -computername $Computer_Line -quiet -count 1
                                If($Check_Connection -eq $true)
                                    {
                                        Invoke-Command -ComputerName $Computer_Line -ScriptBlock ${Function:Set_Lenovo_BIOS_Settings} -ArgumentList (,$CSV_Content), $MyPassword        
                                    }
                                Else
                                    {
                                        write-host "Connexion KO on $Computer_Line" -Foreground yellow                                  
                                    }    
                            }
                    }
                ElseIf(($Computer -ne ""))
                    {
                        Invoke-Command -ComputerName $Computer -ScriptBlock ${Function:Set_Lenovo_BIOS_Settings} -ArgumentList (,$CSV_Content), $MyPassword                                                       
                    }
                Else
                    {
                        Set_Lenovo_BIOS_Settings $CSV_Content $MyPassword    
                    }    
           }
       }
    }
    End
    {
    }
}  
