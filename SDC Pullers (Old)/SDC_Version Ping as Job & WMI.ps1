﻿#https://www.compart.com/en/unicode/U+2717 - Use the HTML Code
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
  Object = [Char]10003
  ForegroundColor = 'Green'
  NoNewLine = $true
  }
$CoolX = @{
  Object = [char]::ConvertFromUtf32(0x1F571)
  ForegroundColor = 'Red'
  NoNewLine = $true
  }
$filepath = ("C:\Temp\SDC_NOT_Pinging.txt")
Clear-Content -Path $filepath -Force
$Ping= @()
$Computers = Read-Host "PC Name"
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
          $Ping | Out-File "C:\Temp\SDC_Pinging.txt"
 }

#================================================================================
Write-Host "===================================================="
Write-Host "=====" -NoNewline -ForegroundColor Black
Write-Host "PING Failures: C:\Temp\SDC_NOT_Pinging.txt" -ForegroundColor DarkCyan -NoNewline
Write-Host "=====" -ForegroundColor Black
Write-Host "===================================================="
Write-Host ""
Write-Host ""
Write-Host "===================================================="
Write-Host "================" -NoNewline -ForegroundColor Black
Write-Host "GETTING SDC VERSIONS" -ForegroundColor Cyan -NoNewline
Write-Host "================" -ForegroundColor Black
Write-Host "===================================================="  
$Array = @()
$58 = '5.8'
$57 = '5.7'
$56 = '5.6'
$55 = '5.5'
$54 = '5.4'
$53 = '5.3'
$PCs = gc "C:\Temp\SDC_Pinging.txt"  
ForEach($Computer in $PCs){ 
                $ErrorActionPreference = 'SilentlyContinue'   
                $OS = (Get-WmiObject -computername $Computer -classname win32_operatingsystem).version
                If($OS -match "​10.0.18362")
                    {
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $58
                        write-host $Computer "is $58" 
                    }
                ElseIf($OS -match "10.0.17763")
                    {
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $57
                        write-host $Computer "is $57" 
                    }
                ElseIf($OS -match "10.0.17134")
                    {
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $56
                        write-host $Computer "is $56"
                    }
                ElseIf($OS -match "10.0.16299")
                    {
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $55
                        write-host $Computer "is $55"
                    }
                ElseIf($OS -match "10.0.15063")
                    {
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $54
                        write-host $Computer "is $54"
                    }  
                ElseIf($OS -match "10.0.14393")
                    {
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $53
                        write-host $Computer "is $53"
                    }else{
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value "WinRm or Permissions Issue"
                        }
                             
            $Array += $obj
            }
$Array | Select ComputerName, OSVersion | OGV -Title "Computer SDCs"
Write-Host "===================================================="
Write-Host "=================" -NoNewline -ForegroundColor Black
Write-Host "Results in Pop-up" -ForegroundColor Cyan -NoNewline
Write-Host "==================" -ForegroundColor Black
Write-Host "===================================================="  


#(Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue).Name
#(Get-CimInstance -ClassName win32_operatingsystem -ErrorAction SilentlyContinue).version