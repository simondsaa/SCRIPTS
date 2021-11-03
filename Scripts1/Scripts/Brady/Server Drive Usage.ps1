$fileName = $null

$LocalUser = Get-WmiObject Win32_ComputerSystem
$LocalEDI = $LocalUser.UserName.TrimStart("AREA52\")
$LocalName = (Get-ADUser "$LocalEDI" -Properties DisplayName).DisplayName

$Date1 = Get-Date -UFormat "%d-%b-%g %H:%M"
$Date2 = Get-Date -Format "dd MMM yy"

$Servers = Get-Content C:\Users\1392134782A\Desktop\Servers.txt

Function Get-DiskSpaceReport
{
    $freeSpaceFileName = "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CS\SCO\SCOO\Server Checks\Disk Space Report $Date2.html"
    
    $Warning = 25
    $Critical = 10

    New-Item -ItemType File $FreeSpaceFileName -Force

    Function writeHtmlHeader
    {
        Param($fileName)
        
        Add-Content $fileName "<html>"
        Add-Content $fileName "<head>"
        Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
        Add-Content $fileName "<title>Disk Space Report</title>"
        add-content $fileName "<style type='text/css'>
        td {
        font-face:tahoma;
        font-size:14px;
        border-top:1px solid #000000;
        border-right:1px solid #000000;
        border-bottom:1px solid #000000;
        border-left:1px solid #000000;
        padding-top:0px;
        padding-right:0px;
        padding-bottom:0px;
        padding-left:0px;
        }
        body {
        margin-left:5px;
        margin-top:5px;
        margin-right:0px;
        margin-bottom:10px;
        }
        table {
        border:thin solid #000000;
        }
        </style>"
        Add-Content $fileName "</head>"
        Add-Content $fileName "<body>"
        add-content $fileName "<table width='100%'>"
        add-content $fileName "<tr bgcolor='#0F6FBB'>"
        add-content $fileName “<td colspan='7' height='25' align='center'>"
        add-content $fileName “<font face='tahoma' color='#FFFFFF' size='4' align='center'><strong>Disk Space Report Generated: </strong>$Date1 by $LocalName</font>"
        add-content $fileName “</td>"
        add-content $fileName “</tr>"
        add-content $fileName “</table>"

    }

    Function writeTableHeader
    {
        Param($fileName)
        Add-Content $fileName "<tr bgcolor=#E7F3FD>"
        Add-Content $fileName "<td width='10%' align='center'>Drive</td>"
        Add-Content $fileName "<td width='50%' align='center'>Drive Label</td>"
        Add-Content $fileName "<td width='10%' align='center'>Total Size (GB)</td>"
        Add-Content $fileName "<td width='10%' align='center'>Used Space (GB)</td>"
        Add-Content $fileName "<td width='10%' align='center'>Freespace (GB)</td>"
        Add-Content $fileName "<td width='10%' align='center'>Freespace %</td>"
        Add-Content $fileName "</tr>"
    }

    Function writeHtmlFooter
    {
        Param($fileName)
        Add-Content $fileName "</body>"
        Add-Content $fileName "</html>"
    }

    Function writeDiskInfo
    {
        Param($fileName,$devId,$volName,$frSpace,$totSpace)
        $totSpace = [Math]::Round(($totSpace/1073741824),0)
        $frSpace = [Math]::Round(($frSpace/1073741824),0)
        $usedSpace = $totSpace – $frspace
        $usedSpace = [Math]::Round($usedSpace,0)
        $freePercent = ($frspace/$totSpace)*100
        $freePercent = [Math]::Round($freePercent,0)
        
        If ($freePercent -gt $Warning)
        {
            Add-Content $fileName "<tr>”
            Add-Content $fileName "<td>$devid</td>"
            Add-Content $fileName "<td>$volName</td>"
            Add-Content $fileName "<td>$totSpace</td>"
            Add-Content $fileName "<td>$usedSpace</td>"
            Add-Content $fileName "<td>$frSpace</td>"
            Add-Content $fileName "<td>$freePercent %</td>"
            Add-Content $fileName "</tr>"
        }
        
        ElseIf ($freePercent -le $Critical)
        {
            Add-Content $fileName "<tr>"
            Add-Content $fileName "<td>$devid</td>"
            Add-Content $fileName "<td>$volName</td>"
            Add-Content $fileName "<td>$totSpace</td>"
            Add-Content $fileName "<td>$usedSpace</td>"
            Add-Content $fileName "<td>$frSpace</td>"
            Add-Content $fileName "<td bgcolor='#FF0000'>$freePercent %</td>"
            Add-Content $fileName "</tr>"
        }
    
        Else
        {
            Add-Content $fileName "<tr>"
            Add-Content $fileName "<td>$devid</td>"
            Add-Content $fileName "<td>$volName</td>"
            Add-Content $fileName "<td>$totSpace</td>"
            Add-Content $fileName "<td>$usedSpace</td>"
            Add-Content $fileName "<td>$frSpace</td>"
            Add-Content $fileName "<td bgcolor='#FBB917'>$freePercent %</td>"
            Add-Content $fileName "</tr>"
        }
    }
    
    writeHtmlHeader $freeSpaceFileName

    ForEach ($Server in $Servers)
    {
        If (Test-Connection -ComputerName $Server -Count 2 -ea 0)
        {
            Add-Content $freeSpaceFileName "<table width='100%'><tbody>"
            Add-Content $freeSpaceFileName "<tr bgcolor='#FFFFE0'>"
            Add-Content $freeSpaceFileName "<td width='100%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong> $Server </strong></font></td>"
            Add-Content $freeSpaceFileName "</tr>"

            writeTableHeader $freeSpaceFileName

            $dp = Get-WmiObject Win32_LogicalDisk -ComputerName $Server -Filter "DriveType=3" -ErrorAction SilentlyContinue
    
            ForEach ($item in $dp)
            {
                #Write-Host $item.DeviceID $item.VolumeName $item.FreeSpace $item.Size
                writeDiskInfo $freeSpaceFileName $item.DeviceID $item.VolumeName $item.FreeSpace $item.Size
            }
        }
    
        Add-Content $freeSpaceFileName “</table>”
    }

    writeHtmlFooter $freeSpaceFileName

    Start-Process -FilePath "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CS\SCO\SCOO\Server Checks\Disk Space Report $Date2.html"
}

Get-DiskSpaceReport