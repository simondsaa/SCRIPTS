#$Path = Read-Host "Path to PCs"
$Output= @()
$Computers = gc "C:\temp\2.txt"
ForEach ($Computer in $Computers)
{
$TC = Test-Connection $Computer -quiet -count 1    
        If($TC -eq $true)
            {
                $Output+= "$Computer"
                Write-host $Computer "is pinging. Attempting to get SDC version." -ForegroundColor Green
            }
        Else
            {
                Write-Host $Computer "is NOT pinging" -ForegroundColor Yellow
            }
      $Output | Out-file "C:\Temp\SDC_Pinging.txt"
 }
#================================================================================
Write-Host "=======================================================" -ForegroundColor Gray 
$Array = @()
$58 = '5.8'
$57 = '5.7'
$56 = '5.6'
$55 = '5.5'
$54 = '5.4'
$53 = '5.3'
$PCs = gc "C:\Temp\SDC_Pinging.txt"  
ForEach($Computer in $PCs){    
            Try {
                $OS = (Get-CimInstance -ComputerName $computer -ClassName win32_operatingsystem -ErrorAction SilentlyContinue).version 
                If($OS -match "​10.0.18362")
                    {
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $58
                        write-host $Computer "is $58" 
                    }
                If($OS -match "10.0.17763")
                    {
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $57
                        write-host $Computer "is $57" 
                    }
                If($OS -match "10.0.17134")
                    {
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $56
                        write-host $Computer "is $56"
                    }
                If($OS -match "10.0.16299")
                    {
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $55
                        write-host $Computer "is $55"
                    }
                If($OS -match "10.0.15063")
                    {
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $54
                        write-host $Computer "is $54"
                    }  
                If($OS -match "10.0.14393")
                    {
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $53
                        write-host $Computer "is $53"
                    }
                    }
                catch
                    {
                        $obj = New-Object PSObject
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                        $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value "WinRm or Permissions Issue"
                    }
                             
            $Array += $obj
            }
$Array | Select ComputerName, OSVersion | OGV -Title "Computer SDCs"


#(Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue).Name
#(Get-CimInstance -ClassName win32_operatingsystem -ErrorAction SilentlyContinue).version 