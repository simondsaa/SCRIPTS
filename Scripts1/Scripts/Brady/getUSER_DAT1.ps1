$erroractionpreference = "SilentlyContinue"

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Machine Name"
$c.Cells.Item(1,2) = "IP Address"
$c.Cells.Item(1,3) = "Last User"
$c.Cells.Item(1,4) = "AV Product"
$c.Cells.Item(1,5) = "Version"
$c.Cells.Item(1,6) = "Scan Engine"
$c.Cells.Item(1,7) = "Virus Definition"
$c.Cells.Item(1,8) = "Virus Definition Date"
$c.Cells.Item(1,9) = "Repository Server"
$c.Cells.Item(1,10) = "Report Time Stamp"
$c.Cells.Item(1,12) = "Systems Offline"


$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

$colComputers = get-content "C:\Users\Timothy.Brady\Desktop\Comps.txt"

foreach ($strComputer in $colComputers)
{
    If (Test-Connection $strComputer -quiet -count 1)
        {
            $c.Cells.Item($intRow,1) = $strComputer

            Function GetProductInfo
            {
                $key="SOFTWARE\McAfee\DesktopProtection"
                $regkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $strComputer)
                $regKey = $regKey.OpenSubKey($key)

                $Product = $regKey.GetValue("Product")
                $c.Cells.Item($intRow,4) = $Product

                $productver = $regKey.GetValue("szProductVer")
                $c.Cells.Item($intRow,5) = $Productver
            }       
            GetProductInfo
                
            Function GetDATInfo
            {
                $key="SOFTWARE\McAfee\AVEngine"
                $regkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $strComputer)
                $regKey = $regKey.OpenSubKey($key)

                $ScanEngine = $regKey.GetValue("EngineVersionMajor")
                $c.Cells.Item($intRow,6) = $ScanEngine

                $VirDefVer = $regKey.GetValue("AVDatVersion")
                $c.Cells.Item($intRow,7) = $VirDefVer

                $virDefDate = $regKey.GetValue("AVDatDate")
                $c.Cells.Item($intRow,8) = $virDefDate
            }
            GetDATInfo
               
            Function GetUserInfo
            {
                 $key="SOFTWARE\Network Associates\ePolicy Orchestrator\Agent"
                 $regkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $strComputer)
                 $regKey = $regKey.OpenSubKey($key)
         
                 $EDIPI = $regKey.GetValue("LoggedOnUser")
                 $ADInfo = Get-ADUser -LDAPFilter "(sAMAccountName=$EDIPI)" -SearchBase "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
                 $Username = $ADInfo.displayName

                 $c.Cells.Item($intRow,3) = $Username
                   
            }        
            GetUserInfo

            Function GetSiteInfo
            {
                $key="SOFTWARE\Network Associates\ePolicy Orchestrator\Agent"
                $regkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$strComputer)
                $regKey = $regKey.OpenSubKey($key)
                $ePOServer = $regKey.GetValue("ePOServerList")
                If ($ePOServer -like "*52MUHJ-HBIA-009*")
                    {
                        $c.Cells.Item($intRow,9) = "True"
                    }
                    Else 
                    {
                        $c.Cells.Item($intRow,9) = "False"
                    }  
            }
            GetSiteInfo
                
            Function GetIPInfo
            {
                 $key="SOFTWARE\McAfee\SystemCore\VSCore\Alert Client\VSE"
                 $regkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $strComputer)
                 $regKey = $regKey.OpenSubKey($key)
         
                 $IPAddress = $regKey.GetValue("Ip")
             
                 $c.Cells.Item($intRow,2) = $IPAddress               
            }        
            GetIPInfo

        $c.Cells.Item($intRow,10) = Get-date

        $intRow = $intRow + 1
        }    
    Else {$c.Cells.Item($intRow,12) = "$strComputer Not Reachable"} 
}
$d.EntireColumn.AutoFit()

$b.SaveAs("\\XLWU-FS-004\325 CS$\SCO\SCOO\AV DAT info.xlsx")
$b.Close()

$a.Quit()