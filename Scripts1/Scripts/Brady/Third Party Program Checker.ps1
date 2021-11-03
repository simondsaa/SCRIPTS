#----------------------------------------------------------------------------------
#                           Written by SrA Timothy Brady
#                           Tyndall AFB, Panama City, FL
#                             Created July 11, 2014
#----------------------------------------------------------------------------------

#Modify these to the latest version available
$FlashVer = "14.0.0.145"
$ShockVer = "12.1.0.150"
$AcrobatVer = "11.0.07"
$JavaVer = "7.0.550"
$IBMVer = "4.0.0.3"

#$Computers = Get-Content "C:\work\test.txt"
$Computer = Read-Host "Computer"
#ForEach ($Computer in $Computers)
#{
    $Progs = @()
    If (Test-Connection $Computer -quiet -BufferSize 16 -Ea 0 -count 1)
    {
        $OS = Get-WmiObject Win32_OperatingSystem -cn $Computer
        If ($OS.OSArchitecture -eq "64-bit"){$RegPath = "Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"}
        ElseIf ($OS.OSArchitecture -eq "32-bit"){$RegPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall"}        
        
        $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer)
        $RegKey = $Reg.OpenSubKey($RegPath)
        $SubKeys = $RegKey.GetSubKeyNames()
        ForEach($Key in $SubKeys)
        {
            $ThisKey = $RegPath+"\\"+$Key 
            $ThisSubKey = $Reg.OpenSubKey($ThisKey)
            $obj = New-Object PSObject
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Name" -Value $($thisSubKey.GetValue("DisplayName"))
            $obj | Add-Member -Force -MemberType NoteProperty -Name "Version" -Value $($thisSubKey.GetValue("DisplayVersion"))
            $Progs += $obj
        }
    }
    Else
    {
        Write-Host "$Computer is not available"
    }

    $Flash = $Progs | Where-Object {$_.Name -like "Adobe Flash*"}
    $Shockwave = $Progs | Where-Object {$_.Name -like "Adobe Shockwave*"}
    $Acrobat = $Progs | Where-Object {$_.Name -like "*Adobe Acrobat*"}
    $Java = $Progs | Where-Object {$_.Name -like "Java*"}
    $IBM = $Progs | Where-Object {$_.Name -like "IBM*"}

    Write-Host
    Write-Host "$Computer"
    If ($Flash.Version -ge "$FlashVer")
    {
        $Flash
        Write-Host -ForegroundColor Green "Flash is up to date"
    }
    Else
    {
        $Flash
        Write-Host -ForegroundColor Red "Flash needs to be updated"
    }
    If ($Shockwave.Version -ge "$ShockVer")
    {
        Write-Host
        $Shockwave
        Write-Host -ForegroundColor Green "Shockwave is up to date"
    }
    Else
    {
        Write-Host
        $Shockwave
        Write-Host -ForegroundColor Red "Shockwave needs to be updated"
    }
    If ($Acrobat.Version -ge "$AcrobatVer")
    {
        Write-Host
        $Acrobat
        Write-Host -ForegroundColor Green "Acrobat is up to date"
    }
    Else
    {
        Write-Host
        $Acrobat
        Write-Host -ForegroundColor Red "Acrobat needs to be updated"
    }
    If ($Java.Version -ge "$JavaVer")
    {
        Write-Host
        $Java
        Write-Host -ForegroundColor Green "Java is up to date"
    }
    Else
    {
        Write-Host
        $Java
        Write-Host -ForegroundColor Red "Java needs to be updated"
    }
    If ($IBM.Version -ge "$IBMVer")
    {
        Write-Host
        $IBM
        Write-Host -ForegroundColor Green "IBM is up to date"
    }
    Else
    {
        Write-Host
        $IBM
        Write-Host -ForegroundColor Red "IBM needs to be updated"
    }
#}