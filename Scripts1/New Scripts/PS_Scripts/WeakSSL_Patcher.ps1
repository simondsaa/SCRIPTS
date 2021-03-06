$LogPath = "C:\Users\timothy.brady\Documents\WeakSSL\Logs"

#====================================================================================
   
Function CheckSSL
{
    Write-Host " "
    Write-Host "Selected 1 WeakSSL Check"
    Write-Host " "

    $Path = Read-Host "Enter path to text file containing list of computer names"

    $Computers = Get-Content $Path

    ForEach($Computer in $Computers)
    {

        $Conn = Test-Connection $Computer -count 1 -quiet

        If($Conn -eq "True")
        {
            $Value = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 128/128').GetValue('Enabled')
            If($Value -eq 1)
            {
                Write-Host -backgroundcolor green -foregroundcolor black "$Computer has been patched"
                Out-File "$LogPath\Has_Patch.txt" -inputobject $Computer -append
            }
            If($value -ne 1)
            { 
                Write-host -backgroundcolor red -foregroundcolor white "$Computer is vulnerable or access is denied"
                Out-File "$LogPath\Vulnerable_Host.txt" -inputobject $Computer -append
            }
        }
        If($Conn -ne "True")
        {
            Write-Host "$Computer offline"
        }
        $Value = $Null
    }
}

#====================================================================================

Function PatchSSL
{
    Write-Host " "
    Write-Host "Selected 2 WeakSSL Patch"
    Write-Host " "

    $Path = Read-Host "Enter path to text file containing list of computer names"

    $Computers = Get-Content $Path

    ForEach($Computer in $Computers)
    {
        Write-Host " " 

        If(Test-Connection -computername $Computer -count 1)
        {
            Write-Host "$Computer is ONLINE"
    
            Try
            {
                $G = $OpenHK = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(‘LocalMachine’,$computer)
                Write-Host "$OpenHK is OPEN"
            }
            Catch
            {
                Write-Host "Can't open HKLM"
            }

            Try
            {
                $RC4val = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$computer).OpenSubKey('SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 128/128').GetValue('Enabled')
                Write-Host "RC4 128/128 Enabled $RC4Val"
            }
            Catch
            {
                Write-Host "Key does not exist"
            }
   
            Try
            {
                $P = $OpenHK.OpenSubKey('SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers',$True)
                Write-Host "Subkey opened for WRITE"
            }
            Catch
            {
                Write-Host "Cant open subkey to write"
            }

            Try
            {
                $I = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(‘LocalMachine’,$Computer).OpenSubKey('SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers',$True).createsubkey('RC4 128/128')
                $I.Setvalue('Enabled’, ‘00000001’, ‘DWORD’)
                Write-Host "Key created"
                Out-File "$LogPath\Success.txt" -inputobject $Computer -append    
            }
            Catch
            {
                Write-Host "Cant write key"
                Out-File "$LogPath\Patch_Failed.txt" -inputobject $Computer -append
            }

            $Counter++
            Write-Host "Count $Counter"
            Write-Host ""

        }
        Else
        {
            Write-Host "$Computer is OFFLINE"
            Out-File "$LogPath\No_Ping.txt" -append -inputobject $Computer
        }
    }    
}

#====================================================================================

Function DeleteSSL
{
    Write-Host " "
    Write-Host "Selected 3 WeakSSL Patch Delete"
    Write-Host " "

    $Path = Read-Host "Enter path to text file containing list of computer names"

    $Computers = Get-Content $Path

    ForEach($Computer in $Computers)
    {
        If(Test-Connection -computername $Computer -count 1 -quiet)
        {
            Write-Host "$Computer is connected"

            $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 128/128').GetValue('Enabled')
   
            If($Reg -eq 1)
            {
                $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(‘LocalMachine’,$Computer)
                $SK = $Reg.OpenSubKey("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers",$True)
                $SK.deletesubkey("RC4 128/128")
            }
        
            Try
            {
                [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers\RC4 128/128').GetValue('Enabled')
            }
            Catch
            {
                Write-Host "Key deleted from $Computer"
            }
        }
    }
}   

#===================================================================================

Do
{
    Write-Host " " 
    Write-Host "=============================================================="
    Write-Host "Weak SSL patch script by Joshua Heitkamp, 502CS, JBSA-Randolph"
    Write-Host "=============================================================="
    Write-Host " " 
    Write-Host "This script must be run as an Administrator"
    Write-Host " "
    Write-Host "Menu"
    Write-Host " "
    Write-Host "1 - Check machines for vulnerability"
    Write-Host "2 - Patch machines"
    Write-Host "3 - Delete patch"
    Write-Host "4 - Clear all logs"
    Write-Host "5 - Exit"
    Write-Host " " 

    $Ans = Read-Host("Make Selection")

    If($Ans -eq 1)
    {
        CheckSSL
    }

    If($Ans -eq 2)
    {
        PatchSSL
    }

    If($Ans -eq 3)
    {
        DeleteSSL
    }

    If($Ans -eq 4)
    {
        $Response = Read-Host "This will delete all logs. Proceed? (Y/N)"
        If($Response -eq "Y")
        {
            Out-File "$LogPath\Has_Patch.txt" -inputobject $null 
            Out-File "$LogPath\No_Ping.txt" -inputobject $null 
            Out-File "$LogPath\Patch_Failed.txt" -inputobject $null 
            Out-File "$LogPath\Success.txt" -inputobject $null 
            Out-File "$LogPath\Vulnerable_Host.txt" -inputobject $null 
        }
   
    }

}
Until($Ans -eq 5)