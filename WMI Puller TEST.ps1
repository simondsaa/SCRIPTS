$Where = Read-Host "PC List"
$Computers = Get-Content $Where
$Path = "C:\Temp\AD Computer Properties.txt"
If (Test-Path $Path){Remove-Item $Path}
ForEach($pcname in $Computers){
        if ($connection = Test-Connection $pcname -Count 1 -Quiet) {
        if ($Bios = Get-WmiObject -class Win32_BIOS -ComputerName $pcname -Filter 'SMBIOSBIOSVERSION like "L06v02.02"') {
            $os = Get-WmiObject -class win32_operatingsystem -ComputerName $pcname
            $comp = Get-WmiObject -class win32_computersystem -ComputerName $pcname
            $java = Get-WmiObject -class Win32_Product -ComputerName $pcname -Filter 'name like "%java%"'
            $vol = Get-WmiObject -class Win32_Volume -ComputerName $pcname
            $net = Get-WmiObject -class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' -ComputerName $pcname
            $CB = Get-WmiObject -class Win32_Product -ComputerName $pcname -Filter 'Name like "%HP%"'
            $Proc = Get-Wmiobject -class Win32_processor -ComputerName $pcname
            
                 New-Object PSObject -Property @{
                'Operating System'    = $os.name.split('|')[0]
                'Version'             = $os.Version
                'Architecture'        = $os.OSArchitecture
                'Serial Number'       = $Bios.SerialNumber
                'Manufacturer'        = $Bios.Manufacturer
                'Bios Version'        = $Bios.SMBIOSBIOSVersion
                'System Name'         = $comp.Name
                'Model'               = $comp.Model
                'Processor'           = $Proc.Name
                'Last Logged In'      = $comp.UserName
                'Java Version'        = $java.name -join ' = '
                'File System'         = $vol.FileSystem
                'IP Address'          = $net.IPAddress -join ' = '
                'Subnet'              = $net.IPSubnet -join ' = '
                'MAC Address'         = $net.MACAddress
                'Carbon Black'        = $CB.name -join ' = '
                'Carbon Black Vendor' = $CB.Vendor -join ' = '
                'Carbon Black Version' = $CB.Version -join ' = '
            } | Select-Object $ServerTable
           
        }
    }




        Write-Output $serverTable | Out-File $Path -append
            }
        
$file = “$Path”
$oXL = New-Object -comobject Excel.Application
$oXL.Visible = $true
$oXL.workbooks.OpenText($file,1,1,1,1,$True,$True,$True,$False,$False,$False)

# 1   Tab = True
# 2   Semicolon = True
# 3   Comma = False
# 4   Space = False
# 5   Other = False

#  C:\Temp\4.txt