#https://www.compart.com/en/unicode/U+2717 - Use the HTML Code
# OR when it's "too long"
#https://stackoverflow.com/questions/1056692/how-do-i-encode-unicode-character-codes-in-a-powershell-string-literal
#$Path = Read-Host "Path to PCs"
#5.5 report 20200925.txt
Write-Host ""
Write-Host "===================================================="
Write-Host "=================" -NoNewline -ForegroundColor Black
Write-Host "STARTING PING TEST" -ForegroundColor DarkCyan -NoNewline
Write-Host "=================" -ForegroundColor Black
Write-Host "===================================================="
$greenCheck = @{
  Object = [Char]::ConvertFromUtf32(0x263B)
  ForegroundColor = 'Green'
  NoNewLine = $true
  }
$CoolX = @{
  Object = [char]::ConvertFromUtf32(0x1F571)
  ForegroundColor = 'Red'
  NoNewLine = $true
  }
$filepath = ("C:\Temp\BIOS\BIOS_NOT_Pinging.txt")
Clear-Content -Path $filepath -Force
$Ping= @()
#$Path = Read-Host "Path to PCs"
$Computers = gc "C:\temp\2.txt"
ForEach ($Computer in $Computers)
{
$TC = Test-Connection $Computer -quiet -count 1   
        If($TC -eq $True)
            {
                $Ping+= "$Computer"
                Write-host $Computer "is ... " -NoNewline
                Write-Host @greenCheck 
                Write-Host ""
            }
        If($TC -eq $False)
            {
                Write-host "$Computer is ... "  -NoNewline              
                Write-Host @Coolx 
                Write-Host ""
                Write-Output "$Computer" >> $filepath
            }  
          $Ping | Out-File "C:\Temp\BIOS\BIOS_Pinging.txt"
 }

#================================================================================
Write-Host "===================================================="
Write-Host "=====" -NoNewline -ForegroundColor Black
Write-Host "PING Failures: C:\Temp\BIOS\BIOS_NOT_Pinging.txt" -ForegroundColor DarkCyan -NoNewline
Write-Host "=====" -ForegroundColor Black
Write-Host "===================================================="
Write-Host ""
Write-Host ""
Write-Host "===================================================="
Write-Host "================" -NoNewline -ForegroundColor Black
Write-Host "GETTING BIOS SETTINGS" -ForegroundColor Cyan -NoNewline
Write-Host "================" -ForegroundColor Black
Write-Host "===================================================="  
$Array = @()
$58 = '5.8'
$57 = '5.7'
$56 = '5.6'
$55 = '5.5'
$54 = '5.4'
$53 = '5.3'
$PCs = gc "C:\Temp\BIOS\BIOS_Pinging.txt"  
ForEach($Computer in $PCs){               
    $OS = (Get-CimInstance -ComputerName $computer -ClassName win32_operatingsystem -ErrorAction SilentlyContinue).version 
    If($OS -match "​10.0.18362")
        {
            
            $obj = New-Object PSObject
            $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
            $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $58
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Setting" -Value $NameOfSetting
            $obj | Add-Member -Force -MemberType NoteProperty -Name "CurrentValue" -Value $CurrentValue
            write-host $Computer "is $58" 
        }
    ElseIf($OS -match "10.0.17763")
        {
            $1 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'Configure Legacy Support and Secure Boot'} | Select-Object -ExpandProperty Name}
            $2 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'Configure Legacy Support and Secure Boot'} | Select-Object -ExpandProperty CurrentValue}
            $3 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'SecureBoot'} | Select-Object -ExpandProperty Name}
            $4 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'SecureBoot'} | Select-Object -ExpandProperty CurrentValue}
            $5 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'SecureBoot'} | Select-Object -ExpandProperty Name}
            $6 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'SecureBoot'} | Select-Object -ExpandProperty CurrentValue}
            $obj = New-Object PSObject
            $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
            $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $57
            $obj | Add-Member -Force -MemberType NoteProperty -Name "P78_Setting" -Value $1
            $obj | Add-Member -Force -MemberType NoteProperty -Name "P78_Value" -Value $2
            $obj | Add-Member -Force -MemberType NoteProperty -Name "L77_Setting" -Value $3
            $obj | Add-Member -Force -MemberType NoteProperty -Name "L77_Value" -Value $4
            $obj | Add-Member -Force -MemberType NoteProperty -Name "L06_Setting" -Value $5
            $obj | Add-Member -Force -MemberType NoteProperty -Name "L06_Value" -Value $6
            write-host $Computer "is $57"  
        }
    ElseIf($OS -match "10.0.17134")
        {
            $obj = New-Object PSObject
            $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
            $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $56
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Setting" -Value $NameOfSetting
            $obj | Add-Member -Force -MemberType NoteProperty -Name "CurrentValue" -Value $CurrentValue
            write-host $Computer "is $56"
        }
    ElseIf($OS -match "10.0.16299")
        {                  
            $1 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'Configure Legacy Support and Secure Boot'} | Select-Object -ExpandProperty Name}
            $2 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'Configure Legacy Support and Secure Boot'} | Select-Object -ExpandProperty CurrentValue}
            $3 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'SecureBoot'} | Select-Object -ExpandProperty Name}
            $4 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'SecureBoot'} | Select-Object -ExpandProperty CurrentValue}
            $5 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'Secure Boot'} | Select-Object -ExpandProperty Name}
            $6 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'Secure Boot'} | Select-Object -ExpandProperty CurrentValue}
            $7 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'UEFI Boot Options'} | Select-Object -ExpandProperty Name}
            $8 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'UEFI Boot Options'} | Select-Object -ExpandProperty CurrentValue}
            $9 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'Virtualization Technology for Directed I/O (VTd)'} | Select-Object -ExpandProperty Name}
            $10 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'Virtualization Technology for Directed I/O (VTd)'} | Select-Object -ExpandProperty CurrentValue}          
            $11 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'Virtualization Technology (VTx)'} | Select-Object -ExpandProperty Name}
            $12 = Invoke-Command -ComputerName $Computer {Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biossetting | Where-Object {$_.Name -eq 'Virtualization Technology (VTx)'} | Select-Object -ExpandProperty CurrentValue}
            $obj = New-Object PSObject
            $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
            $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $55
            $obj | Add-Member -Force -MemberType NoteProperty -Name "P78_SecureBoot-Setting" -Value $1
            $obj | Add-Member -Force -MemberType NoteProperty -Name "P78_SecureBoot-Value" -Value $2
            $obj | Add-Member -Force -MemberType NoteProperty -Name "L77_SecureBoot-Setting" -Value $3
            $obj | Add-Member -Force -MemberType NoteProperty -Name "L77_SecureBoot-Value" -Value $4
            $obj | Add-Member -Force -MemberType NoteProperty -Name "L06_SecureBoot-Setting" -Value $5
            $obj | Add-Member -Force -MemberType NoteProperty -Name "L06_SecureBoot-Value" -Value $6
            $obj | Add-Member -Force -MemberType NoteProperty -Name "P78_UEFI-Setting" -Value $7
            $obj | Add-Member -Force -MemberType NoteProperty -Name "P78_UEFI-Value" -Value $8
            $obj | Add-Member -Force -MemberType NoteProperty -Name "P78_VTd-Setting" -Value $9
            $obj | Add-Member -Force -MemberType NoteProperty -Name "P78_VTd-Value" -Value $10
            $obj | Add-Member -Force -MemberType NoteProperty -Name "P78/L77_VTx-Setting" -Value $11
            $obj | Add-Member -Force -MemberType NoteProperty -Name "P78/L77_VTx-Value" -Value $12
            write-host $Computer "is $55"
        }
    ElseIf($OS -match "10.0.15063")
        {
            $obj = New-Object PSObjectcopy. 
            $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
            $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $54
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Setting" -Value $NameOfSetting
            $obj | Add-Member -Force -MemberType NoteProperty -Name "CurrentValue" -Value $CurrentValue
            write-host $Computer "is $54"
        }  
    ElseIf($OS -match "10.0.14393")
        {
            $obj = New-Object PSObject
            $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
            $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $53
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Setting" -Value $NameOfSetting
            $obj | Add-Member -Force -MemberType NoteProperty -Name "CurrentValue" -Value $CurrentValue
            write-host $Computer "is $53"
        }else{
            $obj = New-Object PSObject
            $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
            $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value "WinRm or Permissions Issue"
            }
                             
$Array += $obj
}
$Array | Select ComputerName, OSVersion, P78_SecureBoot-Setting, P78_SecureBoot-Value, P78_UEFI-Setting, P78_UEFI-Value, P78/L77_VTx-Setting, P78/L77_VTx-Value, P78_VTd-Setting, P78_VTd-Value, L77_SecureBoot-Setting, L77_SecureBoot-Value, L06_SecureBoot-Setting, L06_SecureBoot-Value | OGV -Title "BIOS Settings"
Write-Host "===================================================="
Write-Host "=================" -NoNewline -ForegroundColor Black
Write-Host "Results in Pop-up" -ForegroundColor Cyan -NoNewline
Write-Host "==================" -ForegroundColor Black
Write-Host "===================================================="  


#(Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue).Name
#(Get-CimInstance -ClassName win32_operatingsystem -ErrorAction SilentlyContinue).version 