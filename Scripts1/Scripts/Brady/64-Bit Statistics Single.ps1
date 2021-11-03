$Computer = "TYNAMXSWK001003"

If (Test-Connection $Computer -Quiet -Count 1 -BufferSize 16 -Ea 0)
{
    $Ping = "Online"
    Try
    {
        $Bit = (Get-WmiObject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue).OSArchitecture
        $RAM = [Math]::Round((Get-WmiObject Win32_ComputerSystem -cn $Computer -ErrorAction SilentlyContinue).TotalPhysicalMemory/1048576, 0)
        $Disk = Get-WmiObject Win32_LogicalDisk -cn $Computer -ErrorAction SilentlyContinue | Where {$_.DeviceID -like "C:"}
        $FreeSpace = [Math]::Round($Disk.FreeSpace/1073741824, 0)
    }
    Catch { }
    $Profiles = (Get-ChildItem "\\$Computer\C$\Users").Count
    $AD = Get-ADComputer -Identity $Computer -Properties location, o
    $Org = $AD.o
    $Bldg = ($AD.location).Split(";")[0]
    $Room = ($AD.location).Split(";")[1]
}
Else
{
    $Ping = "Offline"
}

Write-Host $Computer
Write-Host $Ping
Write-Host $Bit
Write-Host $RAM
Write-Host $FreeSpace
Write-Host $Org
Write-Host $Bldg
Write-Host $Room
Write-Host $Profiles