
$computers = Get-Content \\xlwu-fs-05pv\Tyndall_PUBLIC\Patching\Java\JavaPatch.txt    
$sourcefile_x64 = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Applications\Java\Java_8.91\x64\*"
$sourcefile_x86 = "\\xlwu-fs-05pv\Tyndall_PUBLIC\Applications\Java\Java_8.91\x86\*"


#This section will install the software 

foreach ($computer in $computers) 
{   
     
    $destinationFolder_x64 = "\\$computer\C$\TEMP\Java_x64"
    $destinationFolder_x86 = "\\$computer\C$\TEMP\Java_x86"
    
    $App = Get-WmiObject Win32_Product | Where {$_.Name -like "*Java*"}
    $App.Uninstall()
    
    If (Test-Connection -cn $computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        $OSInfo = Get-Wmiobject Win32_OperatingSystem -ComputerName $Computer -ErrorAction SilentlyContinue

        If ($OSInfo.OSArchitecture -eq "64-Bit")
        {
            If (!(Test-Path -path $destinationFolder_x64))
            {
                New-Item $destinationFolder_x64 -Type Directory -Force
            }
            If (!(Test-Path -path $destinationFolder_x86))
            {
                New-Item $destinationFolder_x86 -Type Directory -Force
            }

            Copy-Item -Path $sourcefile_x64 -Destination $destinationFolder_x64
            Copy-Item -Path $sourcefile_x86 -Destination $destinationFolder_x86
            Invoke-Command -ComputerName $computer -ScriptBlock {msiexec.exe /i C:\TEMP\Java_x64\jre1.8.0_91_x64.msi}
            Invoke-Command -ComputerName $computer -ScriptBlock {msiexec.exe /i C:\TEMP\Java_x86\jre1.8.0_91_x86.msi}
            
            #Invoke-Command -ComputerName $computer -ScriptBlock { & cmd /c "msiexec.exe /i c:\TEMP\Java_x86\jre1.8.9_71.msi /qn ADVANCED_OPTIONS=1 CHANNEL=100}
        }

        If ($OSInfo.OSArchitecture -eq "32-Bit")
        {
            If (!(Test-Path -path $destinationFolder_x86))
            {
                New-Item $destinationFolder_x86 -Type Directory -Force
            }

            Copy-Item -Path $sourcefile_x86 -Destination $destinationFolder_x86
            Invoke-Command -ComputerName $computer -ScriptBlock {msiexec.exe /i C:\TEMP\Java_x86\jre1.8.9_71.msi}
        }
    }
    Else
    {
        Write-Host "$computer offline"
    }
}