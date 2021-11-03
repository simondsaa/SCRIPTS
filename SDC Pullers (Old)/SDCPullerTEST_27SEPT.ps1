#$Path = Read-Host "Path to PCs"
$Computers = gc "C:\temp\2.txt"
$Array = @()
ForEach ($Computer in $Computers)
    {   
        $TC = Test-Connection $Computer -quiet -count 1                     
        If($TC)
            {
                $58 = '5.8'
                $57 = '5.7'
                $56 = '5.6'
                $55 = '5.5'
                $54 = '5.4'
                $53 = '5.3'
                $SDC58 = $58 
                $SDC57 = $57
                $SDC56 = $56
                $SDC55 = $55
                $SDC54 = $54
                $SDC53 = $53
                $OS = (Get-CimInstance -ComputerName $computer -ClassName win32_operatingsystem -ErrorAction SilentlyContinue).version            
                If($OS -match "​10.0.18362")
                    {
                        write-host $Computer "is $58"
                $obj = New-Object PSObject
                $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $SDC58 
                    }
                    else{
                If($OS -match "10.0.17763")
                    {
                        write-host $Computer "is $57" 
                $obj = New-Object PSObject
                $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $SDC57 
                    }
                If($OS -match "10.0.17134")
                    {
                        write-host $Computer "is $56"
                $obj = New-Object PSObject
                $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $SDC56 
                    }
                If($OS -match "10.0.16299")
                    {
                        write-host $Computer "is $55"
                $obj = New-Object PSObject
                $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $SDC55 
                    }
                If($OS -match "10.0.15063")
                    {
                        write-host $Computer "is $54"
                $obj = New-Object PSObject
                $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $SDC54 
                    }  
                If($OS -match "10.0.14393")
                    {
                        write-host $Computer "is $53"
                $obj = New-Object PSObject
                $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Computer
                $obj | Add-Member -Force -MemberType NoteProperty -Name "OSVersion" -Value $SDC53 
                    }
                    }
               }                                
            $Array += $obj
    }

$Array | Select ComputerName, OSVersion | OGV -Title "Computer SDCs"

#(Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue).Name
#(Get-CimInstance -ClassName win32_operatingsystem -ErrorAction SilentlyContinue).version 