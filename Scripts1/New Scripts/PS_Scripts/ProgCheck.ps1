$Acrobat = "{AC76BA86-1033-FFFF-7760*"
$Java = "{26A24AE4-039D-4CA4-87B4*"
$Flash = "{A4488E5C-1022-432A-8066*", "{A580818A-6519-4120-AB1C*"
$Shockwave = "{E38C529D-DD73-4002-8489*"
$All = "{AC76BA86-1033-FFFF-7760*", "{A4488E5C-1022-432A-8066*", "{A580818A-6519-4120-AB1C*", "{26A24AE4-039D-4CA4-87B4*", "{E38C529D-DD73-4002-8489*"

Function GetFile
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $Open = New-Object System.Windows.Forms.OpenFileDialog
    $Open.InitialDirectory = "C:\Users"
    $Open.Filter = "All files (*.*)| *.*"
    $Open.ShowDialog() | Out-Null
    $Open.FileName
}

Function ProgCheck
{
    If (Test-Path -Path "C:\Users\1392134782A\Documents\ProgCheck.txt")
    {
        Remove-Item -Path "C:\Users\1392134782A\Documents\ProgCheck.txt" -Force
    }
    "Computer; Program; Version" | Out-File "C:\Users\1392134782A\Documents\ProgCheck.txt" -Force
    
    ForEach ($Computer in $Computers)
    {
        If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
        {
            Try
            {
                $OSInfo = Get-Wmiobject Win32_OperatingSystem -ComputerName $Computer -ErrorAction SilentlyContinue
                If ($OSInfo.OSArchitecture -eq "64-bit"){$RegPath = "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"}
                ElseIf ($OSInfo.OSArchitecture -eq "32-bit"){$RegPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall"}        
                $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer)
                $RegKey = $Reg.OpenSubKey($RegPath)
                $SubKeys = $RegKey.GetSubKeyNames()
                ForEach($Key in $SubKeys)
                {
                    ForEach ($Identifer in $Identifers)
                    {
                        If ($Key -like $Identifer)
                        {
                            $ThisKey = $RegPath+"\"+$Key 
                            $ThisSubKey = $Reg.OpenSubKey($ThisKey)
                            $Display = $ThisSubKey.GetValue("DisplayName")
                            $Version = $ThisSubKey.GetValue("DisplayVersion")
                            Write-Host "$Computer - $Display - $Version" -ForegroundColor Cyan
                            "$Computer; $Display; $Version" | Out-File "C:\Users\1392134782A\Documents\ProgCheck.txt" -Append -Force
                        }
                    }
                }
            }
            Catch
            {
                "$Computer; Access Denied" | Out-File "C:\Users\1392134782A\Documents\ProgCheck.txt" -Append -Force
            }

        }
        Else
        {
            Write-Host "$Computer offline" -ForegroundColor Yellow
            "$Computer; Offline" | Out-File "C:\Users\1392134782A\Documents\ProgCheck.txt" -Append -Force
        }
        Write-Host
    }
    $File = "C:\Users\1392134782A\Documents\ProgCheck.txt"
    $oXL = New-Object -comobject Excel.Application
    $oXL.Visible = $true
    $oXL.workbooks.OpenText($File,1,1,1,1,$True,$True,$True,$False,$False,$False)
    # 1   Tab = True
    # 2   Semicolon = True
    # 3   Comma = False
    # 4   Space = False
    # 5   Other = False
}

Do
{
    Cls
    Write-Host "Program Checker"
    Write-Host
    Write-Host "1 - Acrobat"
    Write-Host "2 - Flash"
    Write-Host "3 - Java"
    Write-Host "4 - Shockwave"
    Write-Host "5 - All Four"
    Write-Host "6 - List of Computers"
    Write-Host "7 - Exit"
    Write-Host

    $Ans = Read-Host "Make Selection"
    
    If ($Ans -eq 1)
    {
        Write-Host
        $Computers = Read-Host "Computer"
        Write-Host
        $Identifers = $Acrobat
        ProgCheck
        Write-Host
        Pause
    }
    If ($Ans -eq 2)
    {
        Write-Host
        $Computers = Read-Host "Computer"
        Write-Host
        $Identifers = $Flash
        ProgCheck
        Write-Host
        Pause
    }
    If ($Ans -eq 3)
    {
        Write-Host
        $Computers = Read-Host "Computer"
        Write-Host
        $Identifers = $Java
        ProgCheck
        Write-Host
        Pause
    }
    If ($Ans -eq 4)
    {
        Write-Host
        $Computers = Read-Host "Computer"
        Write-Host
        $Identifers = $Shockwave
        ProgCheck
        Write-Host
        Pause
    }
    If ($Ans -eq 5)
    {
        Write-Host
        $Computers = Read-Host "Computer"
        Write-Host
        $Identifers = $All
        ProgCheck
        Write-Host
        Pause
    }
    If ($Ans -eq 6)
    {
        Write-Host
        $File = GetFile
        $Computers = Get-Content $File
        Write-Host
        $Identifers = $All
        ProgCheck
        Write-Host
        Pause
    }
}
Until ($Ans -eq 7)