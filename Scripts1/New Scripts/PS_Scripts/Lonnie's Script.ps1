#-----------------------------------------------------------------------------------------#
#                                   Written by Lonnie Stringer                            #
#                                  Tyndall AFB, Panama City, FL                           #
#-----------------------------------------------------------------------------------------#
$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Machine Name"
$c.Cells.Item(1,2) = "AV Product"
$c.Cells.Item(1,3) = "Version"
$c.Cells.Item(1,4) = "Scan Engine"
$c.Cells.Item(1,5) = "Virus Definition"
$c.Cells.Item(1,6) = "Virus Definition Date"
$c.Cells.Item(1,7) = "Repository Server"
$c.Cells.Item(1,8) = "Report Time Stamp"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$colComputers = get-content "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\_Flash\hosts.txt"

foreach ($strComputer in $colComputers)
{
    $c.Cells.Item($intRow,1) = $strComputer

    Function GetProductInfo
    {
        $key="SOFTWARE\McAfee\DesktopProtection"
        $regkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $strComputer)
        $regKey = $regKey.OpenSubKey($key)

        $Product = $regKey.GetValue("Product")
        $c.Cells.Item($intRow,2) = $Product

        $productver = $regKey.GetValue("szProductVer")
        $c.Cells.Item($intRow,3) = $Productver
    }
    
    GetProductInfo
    
    Function GetDATInfo
    {
        $key="SOFTWARE\McAfee\AVEngine"
        $regkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $strComputer)
        $regKey = $regKey.OpenSubKey($key)

        $ScanEngine = $regKey.GetValue("EngineVersionMajor")
        $c.Cells.Item($intRow,4) = $ScanEngine

        $VirDefVer = $regKey.GetValue("AVDatVersion")
        $c.Cells.Item($intRow,5) = $VirDefVer

        $virDefDate = $regKey.GetValue("AVDatDate")
        $c.Cells.Item($intRow,6) = $virDefDate
    }

   GetDATInfo

    Function GetSiteInfo
    {
        $key="SOFTWARE\Network Associates\ePolicy Orchestrator\Agent"
        $regkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $strComputer)
        $regKey = $regKey.OpenSubKey($key)

        $ePOServer = $regKey.GetValue("ePOServerList")
        If ($ePOServer -like "*52MUHJ-HBIA-009*")
        {
            $c.Cells.Item($intRow,7) = "True"
        }
        Else {$c.Cells.Item($intRow,7) = "False"}  
    }

    GetSiteInfo

    $c.Cells.Item($intRow,8) = Get-date

    $intRow = $intRow + 1

}
$d.EntireColumn.AutoFit()