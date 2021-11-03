        $Path = Read-Host "PCs"
        $ComputerName = Get-Content $Path
        Foreach ($Computer in $ComputerName) {
            try {
                $RemoteRegistry = Get-CimInstance -Class Win32_Service -ComputerName $Computer -Filter 'Name = "RemoteRegistry"' -ErrorAction Stop
                if ($RemoteRegistry.State -eq 'Running') {
                    Write-Output "$Computer is already Enabled"
                }
 
                if ($RemoteRegistry.StartMode -eq 'Disabled') {
                    Set-Service -Name RemoteRegistry -ComputerName $Computer -StartupType Manual -ErrorAction Stop
                    Write-Output "$Computer : Remote Registry has been Enabled"
                }
 
                if ($RemoteRegistry.State -eq 'Stopped') {
                    Start-Service -InputObject (Get-Service -Name RemoteRegistry -ComputerName $Computer) -ErrorAction Stop
                    Write-Output "$Computer : Remote Registry has been Started"
                }
 
            } catch {
                $ErrorMessage = $Computer + " Error: " + $_.Exception.Message
 
            }
        }