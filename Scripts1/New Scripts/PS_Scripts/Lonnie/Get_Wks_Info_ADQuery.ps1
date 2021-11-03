#Creating Excel Spreadsheet

$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = "Machine Name"
$c.Cells.Item(1,2) = "Manufacturer"
$c.Cells.Item(1,3) = "System Model"
$c.Cells.Item(1,4) = "Operating System"
$c.Cells.Item(1,5) = "SDC Version"
$c.Cells.Item(1,6) = "System Arch"
$c.Cells.Item(1,7) = "IP Address"
$c.Cells.Item(1,8) = "MAC Address"
#$c.Cells.Item(1,9) = "RAM"

$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2


$domain = "OU=Tyndall AFB Computers,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$objDomain = [adsi]("LDAP://" + $domain)
$search = New-Object System.DirectoryServices.DirectorySearcher
$search.SearchRoot = $objDomain
$search.Filter = "(&(objectClass=computer))"
$search.SearchScope = "Subtree"
$search.PageSize = 99999
$results = $search.FindAll()

# Edit this to your IP range for your network
$IPRange = "131.55*"

ForEach($computer in $results)
{
    $objComputer = $computer.GetDirectoryEntry()
    $CompName = $objComputer.cn

    If (Test-Connection $CompName -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        
        # Mod 2:
        # ------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	    Try
	    {        
	        $RegOpen = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$CompName)
	        $RegKey = $RegOpen.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation')
	        $SDC = $RegKey.GetValue('Model')
	    }
	    Catch
	    {
	        $SDC = "N/A"
	    }
        # ------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	    $Comp = Get-WmiObject Win32_ComputerSystem -cn $CompName -ErrorAction SilentlyContinue
        $OS = Get-WmiObject Win32_OperatingSystem -cn $CompName | Select-Object Caption -ErrorAction SilentlyContinue
        #$OS_TrimStart = $OS.TrimStart("@").TrimStart("{").TrimStart("Caption").TrimStart("=")
        $RAM = [Math]::Round($Comp.TotalPhysicalMemory/1048576, 0)
	    $AD = Get-ADComputer -LDAPFilter "(name=$CompName)" -Properties whenCreated
        #$NIC = Get-WmiObject Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $CompName | select IPAddress -ErrorAction SilentlyContinue
        $NIC = ([System.Net.Dns]::GetHostByName($CompName).AddressList[0]).IpAddressToString
        #$MAC = Get-WmiObject Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $CompName | select MACAddress -ErrorAction SilentlyContinue
        $IPMAC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $CompName
        $MAC = ($IPMAC | where { $_.IpAddress -eq $NIC}).MACAddress
        $CompMod = Get-WmiObject Win32_ComputerSystem -cn $CompName -ErrorAction SilentlyContinue
        $Man = $CompMod.Manufacturer
        $Model = $CompMod.Model
        $Bit = $CompMod.SystemType
  
        $c.Cells.Item($intRow,1) = "$CompName"
        $c.Cells.Item($intRow,2) = "$Man"
        $c.Cells.Item($intRow,3) = "$Model"
        $c.Cells.Item($intRow,4) = "$OS"
        $c.Cells.Item($intRow,5) = "$SDC"
        $c.Cells.Item($intRow,6) = "$Bit"
        $c.Cells.Item($intRow,7) = "$NIC"
        $c.Cells.Item($intRow,8) = "$MAC"
        #$c.Cells.Item($intRow,9) = "$RAM"

            $intRow = $intRow + 1
        }

    }