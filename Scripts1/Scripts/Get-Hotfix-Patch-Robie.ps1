﻿#### Spreadsheet Location 
 $DirectoryToSaveTo = "C:\Users\1180219788A\Desktop\"
 $date=Get-Date -format "yyyy-MM-dd" 
 $Filename="Patchinfo-$($date)" 
  
###InputLocation 
$Computers = Get-Content C:\Users\1180219788A\Desktop\Computers.txt 
#$Computers = Read-Host "Computer Name"
  
#Create a new Excel object using COM  
$Excel = New-Object -ComObject Excel.Application 
$Excel.visible = $True 
$Excel = $Excel.Workbooks.Add() 
$Sheet = $Excel.Worksheets.Item(1) 
 
$sheet.Name = 'MS17-10 Patch status - ' 
#Create a Title for the first worksheet 
$row = 1 
$Column = 1 
$Sheet.Cells.Item($row,$column)= 'Patch status'  
 
$range = $Sheet.Range("a1","f2") 
$range.Merge() | Out-Null 
$range.VerticalAlignment = -4160 
 
#Give it a nice Style so it stands out 
$range.Style = 'Title' 
 
#Increment row for next set of data 
$row++;$row++ 
 
#Save the initial row so it can be used later to create a border 
#Counter variable for rows 
$intRow = $row 
$xlOpenXMLWorkbook=[int]51 
 
 
$Sheet.Cells.Item($intRow,1)  ="Hostname" 
$Sheet.Cells.Item($intRow,2)  ="Reachable" 
$Sheet.Cells.Item($intRow,3)  ="Patch Status" 
$Sheet.Cells.Item($intRow,4)  ="Operating System" 
$Sheet.Cells.Item($intRow,5)  ="System Type" 
$Sheet.Cells.Item($intRow,6)  ="Last Boot Time" 
 
 
for ($col = 1; $col –le 6; $col++) 
     { 
          $Sheet.Cells.Item($intRow,$col).Font.Bold = $True 
          $Sheet.Cells.Item($intRow,$col).Interior.ColorIndex = 48 
          $Sheet.Cells.Item($intRow,$col).Font.ColorIndex = 34 
     } 
 
$intRow++ 
 
 
Function GetStatusCode 
{  
    Param([int] $StatusCode)   
    switch($StatusCode) 
    { 
        0         {"Success"} 
        11001   {"Buffer Too Small"} 
        11002   {"Destination Net Unreachable"} 
        11003   {"Destination Host Unreachable"} 
        11004   {"Destination Protocol Unreachable"} 
        11005   {"Destination Port Unreachable"} 
        11006   {"No Resources"} 
        11007   {"Bad Option"} 
        11008   {"Hardware Error"} 
        11009   {"Packet Too Big"} 
        11010   {"Request Timed Out"} 
        11011   {"Bad Request"} 
        11012   {"Bad Route"} 
        11013   {"TimeToLive Expired Transit"} 
        11014   {"TimeToLive Expired Reassembly"} 
        11015   {"Parameter Problem"} 
        11016   {"Source Quench"} 
        11017   {"Option Too Big"} 
        11018   {"Bad Destination"} 
        11032   {"Negotiating IPSEC"} 
        11050   {"General Failure"} 
        default {"Failed"} 
    } 
} 
 
 
 
Function GetUpTime 
{ 
    param([string] $LastBootTime) 
    $Uptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime) 
    "Days: $($Uptime.Days); Hours: $($Uptime.Hours); Minutes: $($Uptime.Minutes); Seconds: $($Uptime.Seconds)"  
} 
 
 
foreach ($Computer in $Computers) 
 { 
 
 TRY { 
        $OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer 
        $sheetS = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer 
        $sheetPU = Get-WmiObject -Class Win32_Processor -ComputerName $Computer 
        $drives = Get-WmiObject -ComputerName $Computer Win32_LogicalDisk | Where-Object {$_.DriveType -eq 3} 
        $pingStatus = Get-WmiObject -Query "Select * from win32_PingStatus where Address='$Computer'" 
        $OSRunning = $OS.caption + " " + $OS.OSArchitecture + " SP " + $OS.ServicePackMajorVersion 
        $systemType=$sheetS.SystemType 
        $date = Get-Date 
        $uptime = $OS.ConvertToDateTime($OS.lastbootuptime) 

        If ($OS.caption -like "*Windows 7*")
            
            { $Patch = "KB4019263" }
        
        ElseIf  ($OS.caption -like "*Windows 10*")
            
            { $Patch = "KB4019473" }

        Else 
            { $Patch = "" }
   
 if
   
 ($kb=get-hotfix -id $Patch -ComputerName $computer -ErrorAction 2) 
 
 { 
 $kbinstall="$patch is installed" 
 } 
 else 
 { 
 $kbinstall="$patch is not installed" 
 } 
 
  
  
 if($pingStatus.StatusCode -eq 0) 
    { 
        $Status = GetStatusCode( $pingStatus.StatusCode ) 
    } 
else 
    { 
    $Status = GetStatusCode( $pingStatus.StatusCode ) 
       } 
 } 
  
 CATCH 
 { 
 $pcnotfound = "true" 
 } 
 #### Pump Data to Excel 
 if ($pcnotfound -eq "true") 
 { 
 #$sheet.Cells.Item($intRow, 1) = "PC Not Found" 
 $sheet.Cells.Item($intRow, 1) = $computer 
 $sheet.Cells.Item($intRow, 2) = "Unreachable" 
 } 
 else 
 { 
 $sheet.Cells.Item($intRow, 1) = $computer 
 $sheet.Cells.Item($intRow, 2) = $status 
 $Sheet.Cells.Item($intRow, 3) = $kbinstall 
 $sheet.Cells.Item($intRow, 4) = $OSRunning 
 $Sheet.Cells.Item($intRow, 5) = $SystemType 
 $sheet.Cells.Item($intRow, 6) = $uptime 
 } 
 
  
$intRow = $intRow + 1 
 $pcnotfound = "false" 
 } 
 
$erroractionpreference = “SilentlyContinue”  
 
$Sheet.UsedRange.EntireColumn.AutoFit() 
########################################333 
 
 
 
############################################################## 
 
$filename = "$DirectoryToSaveTo$filename.xlsx" 
#if (test-path $filename ) { rm $filename } #delete the file if it already exists 
$Sheet.UsedRange.EntireColumn.AutoFit() 
$Excel.SaveAs($filename, $xlOpenXMLWorkbook) #save as an XML Workbook (xslx) 
$Excel.Saved = $True 
#$Excel.Close() 
$Excel.DisplayAlerts = $False 
#$Excel.quit() 