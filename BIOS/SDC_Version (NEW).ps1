#$Path = Read-Host "Path to PCs"
$Computers = gc "C:\temp\2.txt"
$Array = @()
$TC = Test-Connection $Computer -quiet -count 1
ForEach ($Computer in $Computers)
    {                        
        If($TC)
            {
                $58 = '5.8'
                $57 = '5.7'
                $56 = '5.6'
                $55 = '5.5'
                $54 = '5.4'
                $53 = '5.3'
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
            $Array += $obj
    }

$Array | Select ComputerName, OSVersion | OGV -Title "Computer SDCs"

#(Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue).Name
#(Get-CimInstance -ClassName win32_operatingsystem -ErrorAction SilentlyContinue).version 