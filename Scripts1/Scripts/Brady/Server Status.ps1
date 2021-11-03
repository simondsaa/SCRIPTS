$Path = "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CS\SCO\SCOO\Server Checks"
$Date = Get-Date -Format "dd MMM yy"

$MonthNum = (Get-Date).Month
$Month = Get-Date -Format "MMM yy"
$Folder = "0$MonthNum. $Month"

If (!(Test-Path -Path "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CS\SCO\SCOO\Server Checks\$Folder"))
{
    New-Item -ItemType Directory -Path "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CS\SCO\SCOO\Server Checks\$Folder" -Force | Out-Null
}

If (Test-Path "$Path\$Folder\Server Status $Date.xlsx")
{
    Remove-Item "$Path\$Folder\Server Status $Date.xlsx" -Force
}

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Server Name"
$c.Cells.Item(1,2) = "Status"
$c.Cells.Item(1,3) = "Reboot Status"
$c.Cells.Item(1,4) = "SCCM Service Status"
$c.Cells.Item(1,5) = "Hard Drive Space"
$c.Cells.Item(1,6) = "Server Uptime"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$WindowsServers = Get-Content "$Path\Windows Servers.txt"
$LinuxServers = Get-Content "$Path\Linux Servers.txt"
$AllServers = $WindowsServers + $LinuxServers

ForEach ($Server in $AllServers)
{
    $c.Cells.Item($intRow,1) = $Server
    
    If (Test-Connection $Server -Count 1 -ea 0)
    {
        $Ping = "Online"
        If ($Server -notlike "*vmh*")
        {
            
            Try
            {
                $OS = Get-WmiObject Win32_OperatingSystem -cn $Server -ErrorAction SilentlyContinue
                $Disks = Get-WmiObject Win32_LogicalDisk -cn $Server -Filter "DriveType=3" -ErrorAction SilentlyContinue
                $Uptime = (Get-Date) – [System.Management.ManagementDateTimeconverter]::ToDateTime($OS.LastBootUpTime)
                $Days = $Uptime.Days
                $Hours = $Uptime.Hours
                $Minutes = $Uptime.Minutes
            } 
            
            Catch
            {
                $Disks = "Access failed"
            }
            
            $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Server)
            $RegPath = "SYSTEM\CurrentControlSet\Control\Session Manager"
            $Key = $Registry.OpenSubKey($RegPath)
            $Value = $Key.GetValue('PendingFileRenameOperations')

            If ($Value -ne $null)
            {
                $Color = "Yellow"
                $Reboot = "Pending reboot"
            }
            Else
            {
                $Color = "White"
                $Reboot = "No reboot required"
            }

            If ((Get-Service -cn $Server -Name CcmExec).Status -eq "Running")
            {
                $SCCM = "Running"
            }
            Else
            {
                $SCCM = "Service not running"
            }

            $UptimeF = "$Days d $Hours h $Minutes m"

            $c.Cells.Item($intRow,2) = $Ping
            $c.Cells.Item($intRow,3) = $Reboot
            $c.Cells.Item($intRow,4) = $SCCM
            $c.Cells.Item($intRow,6) = $UptimeF
                            
            If ($Disks -eq "Access failed")
            {
                $c.Cells.Item($intRow,5) = $Disks

                $intRow = $intRow + 1
            }
            Else
            {
                ForEach ($Disk in $Disks)
                {
                    $DeviceID = $Disk.DeviceID
                    $Size = $Disk.Size
                    $FreeSpace = $Disk.FreeSpace

                    $PercentFree = [Math]::Round(($FreeSpace/$Size) * 100, 0)
                
                    $c.Cells.Item($intRow,5) = "$DeviceID $PercentFree %"

                    $intRow = $intRow + 1
                }
            }
        }

        Else
        {
            $Reboot = "N/A"
            $SCCM = "N/A"
            $UptimeF = "N/A"

            $c.Cells.Item($intRow,2) = $Ping
            $c.Cells.Item($intRow,3) = $Reboot
            $c.Cells.Item($intRow,4) = $SCCM
            $c.Cells.Item($intRow,6) = $UptimeF

            $intRow = $intRow + 1
        }
    }
    Else
    {
        $Ping = "Offline"
        $Reboot = "N/A"
        $SCCM = "N/A"
        $UptimeF = "N/A"

        $c.Cells.Item($intRow,2) = $Ping
        $c.Cells.Item($intRow,3) = $Reboot
        $c.Cells.Item($intRow,4) = $SCCM
        $c.Cells.Item($intRow,6) = $UptimeF

        $intRow = $intRow + 1
    }

    #Write-Host "$Server - $Ping - $Reboot - $SCCM - $UptimeF" -ForegroundColor $Color
}

$d.EntireColumn.AutoFit()

$b.SaveAs("$Path\$Folder\Server Status $Date.xlsx")
$b.Close()

$a.Quit()