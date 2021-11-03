$Path = "C:\Users\1392134782A\Desktop\Server Info.xls"

$IPRange = "131.55*"

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Server Name"
$c.Cells.Item(1,2) = "Server Model"
$c.Cells.Item(1,3) = "IP Address"
$c.Cells.Item(1,4) = "MAC Address"
$c.Cells.Item(1,5) = "Operating System"
$c.Cells.Item(1,6) = "Service Pack"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$Computers = Get-Content C:\Users\1392134782A\Desktop\Servers.txt

ForEach ($Computer in $Computers)
{
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        $Comp = Get-WmiObject Win32_ComputerSystem -cn $Computer -ErrorAction SilentlyContinue
        $OS = Get-Wmiobject Win32_OperatingSystem -cn $Computer -ErrorAction SilentlyContinue
        $NIC = Get-WmiObject Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $Computer -ErrorAction SilentlyContinue |
        Where-Object {$_.IPAddress -like "$IPRange"}
        
        $Model = $Comp.Manufacturer+" "+$Comp.Model
        
        $c.Cells.Item($intRow,1) = $Comp.Name
        $c.Cells.Item($intRow,2) = $Model
        $c.Cells.Item($intRow,3) = $NIC.IPAddress
        $c.Cells.Item($intRow,4) = $NIC.MACAddress
        $c.Cells.Item($intRow,5) = $OS.Caption
        $c.Cells.Item($intRow,6) = "SP"+$OS.ServicePackMajorVersion
    }
    Else
    {
        $c.Cells.Item($intRow,3) = "No ping"
    }

    $intRow = $intRow + 1
}

$d.EntireColumn.AutoFit()

$b.SaveAs($Path)