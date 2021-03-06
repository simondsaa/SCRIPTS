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

$colComputers = get-content "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\HBSS\No_Agent_Targets.txt"

foreach ($strComputer in $colComputers)
{
    If (Test-Connection $strComputer -quiet -count 1)
        {
            $c.Cells.Item($intRow,1) = $strComputer
            
            Function GetIPInfo
            {
                 
                 $NetInfo = Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled = $true" -ComputerName $strComputer -ErrorAction SilentlyContinue | Where-Object {$_.IPAddress -like "131.55*"}
                 $NIC = $NetInfo.Description
                 $IP = $NetInfo.IPAddress
                 $MAC = $NetInfo.MACAddress                 
                        
                 $c.Cells.Item($intRow,2) = $IP               
            }        
            GetIPInfo

            <#

            Function GetUserInfo
            {
                 If((GWmi win32_operatingsystem -computername $strComputer).osarchitecture -eq "32-bit")
                 {
                 $GetEDI=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $strComputer).OpenSubKey('SOFTWARE\Network Associates\ePolicy Orchestrator\Agent').GetValue('LoggedOnUser')       
                 $ADInfo = Get-ADUser -LDAPFilter "(sAMAccountName=$GetEDI)" -SearchBase "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
                 $Username = $ADInfo.displayName
                 Write-Host "$Username"


                 $c.Cells.Item($intRow,3) = $GetEDI
                 }
                 If((GWmi win32_operatingsystem -computername $strComputer).osarchitecture -eq "64-bit")
                 { 
                 $GetEDI=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $strComputer).OpenSubKey('SOFTWARE\WOW6432Node\Network Associates\ePolicy Orchestrator\Agent').GetValue('LoggedOnUser')       
                 $ADInfo = Get-ADUser -LDAPFilter "(sAMAccountName=$GetEDI)" -SearchBase "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
                 $Username = $ADInfo.displayName
                 Write-Host "$Username"

                 $c.Cells.Item($intRow,3) = $GetEDI
                 } 
            }        
            GetUserInfo

            #>
             
            Function Get-McAfeeVersion 
            { 

                If((GWmi win32_operatingsystem -computername $strComputer).osarchitecture -eq "32-bit")
                {
                $GetEDI=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $strComputer).OpenSubKey('SOFTWARE\Network Associates\ePolicy Orchestrator\Agent').GetValue('LoggedOnUser')       
                $Product = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$strComputer).OpenSubKey('SOFTWARE\McAfee\DesktopProtection').GetValue('Product') 
                $ProductVer = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$strComputer).OpenSubKey('SOFTWARE\McAfee\DesktopProtection').GetValue('szProductVer') 
                $EngineVer = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$strComputer).OpenSubKey('SOFTWARE\McAfee\AVEngine').GetValue('EngineVersionMajor') 
                $DatVer = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$strComputer).OpenSubKey('SOFTWARE\McAfee\AVEngine').GetValue('AVDatVersion') 
                $DatDate = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$strComputer).OpenSubKey('SOFTWARE\McAfee\AVEngine').GetValue('AVDatDate') 
                $ePOServer = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$strComputer).OpenSubKey('SOFTWARE\Network Associates\ePolicy Orchestrator\Agent').GetValue('ePOServerList')
                
                $c.Cells.Item($intRow,3) = $GetEDI
                $c.Cells.Item($intRow,4) = $Product
                $c.Cells.Item($intRow,5) = $Productver
                $c.Cells.Item($intRow,6) = $EngineVer
                $c.Cells.Item($intRow,7) = $DatVer                
                $c.Cells.Item($intRow,8) = $DatDate

                    If ($ePOServer -like "*MUHJ-HB*")
                        {
                        $c.Cells.Item($intRow,9) = "True"
                        }
                    Else 
                        {
                        $c.Cells.Item($intRow,9) = "False"
                        }
                $c.Cells.Item($intRow,10) = Get-date

                }
        
                If((GWmi win32_operatingsystem -computername $strComputer).osarchitecture -eq "64-bit")
                {
                $GetEDI=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $strComputer).OpenSubKey('SOFTWARE\WOW6432Node\Network Associates\ePolicy Orchestrator\Agent').GetValue('LoggedOnUser')       
                $Product64 = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$strComputer).OpenSubKey('SOFTWARE\WOW6432Node\McAfee\DesktopProtection').GetValue('Product') 
                $ProductVer64 = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$strComputer).OpenSubKey('SOFTWARE\WOW6432Node\McAfee\DesktopProtection').GetValue('szProductVer') 
                $EngineVer64 = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$strComputer).OpenSubKey('SOFTWARE\WOW6432Node\McAfee\AVEngine').GetValue('EngineVersionMajor') 
                $DatVer64 = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$strComputer).OpenSubKey('SOFTWARE\WOW6432Node\McAfee\AVEngine').GetValue('AVDatVersion') 
                $DatDate64 = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$strComputer).OpenSubKey('SOFTWARE\WOW6432Node\McAfee\AVEngine').GetValue('AVDatDate') 
                $ePOServer64 = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$strComputer).OpenSubKey('SOFTWARE\WOW6432Node\Network Associates\ePolicy Orchestrator\Agent').GetValue('ePOServerList')

                $c.Cells.Item($intRow,3) = $GetEDI
                $c.Cells.Item($intRow,4) = $Product64
                $c.Cells.Item($intRow,5) = $Productver64
                $c.Cells.Item($intRow,6) = $EngineVer64
                $c.Cells.Item($intRow,7) = $DatVer64                
                $c.Cells.Item($intRow,8) = $DatDate64

                    If ($ePOServer64 -like "*MUHJ-HB*")
                        {
                        $c.Cells.Item($intRow,9) = "True"
                        }
                    Else 
                        {
                        $c.Cells.Item($intRow,9) = "False"
                        }
                $c.Cells.Item($intRow,10) = Get-date
                }
            }
            Get-McAfeeVersion             




        $intRow = $intRow + 1
        } 
           
    Else {$c.Cells.Item($intRow,12) = "$strComputer Not Reachable"}
    
    
    <#{$c.Cells.Item($intRow,1) = "$strComputer Not Reachable"} #>
}
$d.EntireColumn.AutoFit()

$b.SaveAs("C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\HBSS\AV DAT info.xlsx")
#$b.Close()

#$a.Quit()