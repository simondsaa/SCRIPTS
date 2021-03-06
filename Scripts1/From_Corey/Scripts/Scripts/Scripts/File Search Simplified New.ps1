Write-Host
Write-Host "----------------------------------------------------------------------------------
                           Written by SrA Timothy Brady
                           Tyndall AFB, Panama City, FL
----------------------------------------------------------------------------------"
$date = Get-Date -format dd-MMM-yyyy
$server = Read-Host "Which File Server Would You Like to Search? 1, 2, 3, or 4"
Write-Host
If ($server -eq "1"){$filepath = "\\XLWU-FS-01\"
Write-Host "Available Subfolders: 
 0 = root
 1 = ANG\1AF
 2 = ANG\601 AOG
 3 = ANG\601 COD
 4 = ANG\601 CPD
 5 = ANG\601 SD
 6 = ANG\CONR
 7 = ANG\COS
 8 = ANG\IMO
 Or type the Path if known"
Write-Host
    $subfile1 = Read-Host "Subfolder *OPTIONAL*"
        If ($subfile1 -eq "0"){($directory = "$filepath"+"root") -and ($OutFile = "FS-001")}
        ElseIf ($subfile1 -eq "1"){($directory = "$filepath"+"root\ANG\1AF") -and ($OutFile = "1 AF")}
        ElseIf ($subfile1 -eq "2"){($directory = "$filepath"+"root\ANG\601 AOG") -and ($OutFile = "601 AOG")}
        ElseIf ($subfile1 -eq "3"){($directory = "$filepath"+"root\ANG\601 COD") -and ($OutFile = "601 COD")}
        ElseIf ($subfile1 -eq "4"){($directory = "$filepath"+"root\ANG\601 CPD") -and ($OutFile = "601 CPD")}
        ElseIf ($subfile1 -eq "5"){($directory = "$filepath"+"root\ANG\601 SD") -and ($OutFile = "601 SD")}
        ElseIf ($subfile1 -eq "6"){($directory = "$filepath"+"root\ANG\CONR") -and ($OutFile = "CONR")}
        ElseIf ($subfile1 -eq "7"){($directory = "$filepath"+"root\ANG\COS") -and ($OutFile = "COS")}
        ElseIf ($subfile1 -eq "8"){($directory = "$filepath"+"root\ANG\IMO") -and ($OutFile = "IMO")}
        ElseIf ($subfile1 -ne "0-8"){($directory = "$filepath"+"$subfile1") -and ($OutFile = "$subfile1")}
        }
ElseIf ($server -eq "2"){$filepath = "\\XLWU-FS-002\"
Write-Host "Available Subfolders:
 0 = root
 1 = NCOA\Shared
 2 = RHS\Shared
 3 = tyndall
 Or type the Path if known"
Write-Host
    $subfile2 = Read-Host "Subfolder *OPTIONAL*"
        If ($subfile2 -eq "0"){($directory = "$filepath"+"root") -and ($OutFile = "FS-002")}
        ElseIf ($subfile2 -eq "1"){($directory = "$filepath"+"root\NCOA\Shared") -and ($OutFile = "NCOA")}
        ElseIf ($subfile2 -eq "2"){($directory = "$filepath"+"root\RHS\Shared") -and ($OutFile = "823 RHS")}
        ElseIf ($subfile2 -eq "3"){($directory = "$filepath"+"root\tyndall") -and ($OutFile = "tyndall")}
        ElseIf ($subfile2 -ne "0-3"){($directory = "$filepath"+"$subfile2") -and ($OutFile = "$subfile2")}
        }
ElseIf ($server -eq "3"){$filepath = "\\XLWU-FS-003\"
Write-Host "Available Subfolders:
 0 = root
 1 = 53WEG\Shared
 2 = 361 TRS\361 TRS
 3 = 479_FTG\451FTS
 4 = 479_FTG\455FTS
 5 = 479_FTG\479FTG
 6 = 479_FTG\479OSS
 7 = AFCESA\Shared
 8 = tyndall
 Or type the Path if known"
Write-Host
    $subfile3 = Read-Host "Subfolder *OPTIONAL*"
        If ($subfile3 -eq "0"){($directory = "$filepath"+"root") -and ($OutFile = "FS-003")}
        ElseIf ($subfile3 -eq "1"){($directory = "$filepath"+"root\53WEG\Shared") -and ($OutFile = "53 WEG")}
        ElseIf ($subfile3 -eq "2"){($directory = "$filepath"+"root\361 TRS\361 TRS") -and ($OutFile = "361 TRS")}
        ElseIf ($subfile3 -eq "3"){($directory = "$filepath"+"root\479_FTG\451FTS") -and ($OutFile = "451 FTS")}
        ElseIf ($subfile3 -eq "4"){($directory = "$filepath"+"root\479_FTG\455FTS") -and ($OutFile = "455 FTS")}
        ElseIf ($subfile3 -eq "5"){($directory = "$filepath"+"root\479_FTG\479FTG") -and ($OutFile = "479 FTG")}
        ElseIf ($subfile3 -eq "6"){($directory = "$filepath"+"root\479_FTG\479OSS") -and ($OutFile = "479 OSS")}
        ElseIf ($subfile3 -eq "7"){($directory = "$filepath"+"root\AFCESA\Shared") -and ($OutFile = "AFCESA")}
        ElseIf ($subfile3 -eq "8"){($directory = "$filepath"+"root\tyndall") -and ($OutFile = "tyndall")}
        ElseIf ($subfile3 -ne "0-8"){($directory = "$filepath"+"$subfile3") -and ($OutFile = "$subfile3")}
        }
ElseIf ($server -eq "4"){$filepath = "\\XLWU-FS-004\"
Write-Host "Available Subfolders:
 0 = root
 1 = 325 FW
 2 = 325 FW Staff
 3 = 325 FW Public
 4 = 325 MSG
 5 = 325 MXG
 6 = 325 OG
 7 = 325 MSG Staff
 8 = 325 MSG Public
 9 = 325 CES
 10 = 325 CONS
 11 = 325 CS
 12 = 325 FSS
 13 = 325 LRD
 14 = 325 SFS
 15 = 325 MXG Staff
 16 = 325 MXG Public
 17 = 325 AMXS
 18 = 325 MOS
 19 = 325 MXS
 20 = 372 TRS
 21 = LOCMAR
 22 = 325 OG Staff
 23 = 325 OG Public
 24 = 43 FS
 25 = 95 FS
 26 = 325 OSS
 27 = 325 TRSS
 Or type the Path if known"
Write-Host
    $subfile4 = Read-Host "Subfolder *OPTIONAL*"
        If ($subfile4 -eq "0"){($directory = "$filepath"+"root") -and ($OutFile = "FS-004")}
        ElseIf ($subfile4 -eq "1"){($directory = "$filepath"+"root\325 FW") -and ($OutFile = "325 FW")}
        ElseIf ($subfile4 -eq "2"){($directory = "$filepath"+"root\325 FW\325 FW Staff") -and ($OutFile = "FW Staff")}
        ElseIf ($subfile4 -eq "3"){($directory = "$filepath"+"root\325 FW\325 FW Public") -and ($OutFile = "FW Public")}
        ElseIf ($subfile4 -eq "4"){($directory = "$filepath"+"root\325 FW\325 MSG") -and ($OutFile = "325 MSG")}
        ElseIf ($subfile4 -eq "5"){($directory = "$filepath"+"root\325 FW\325 MXG") -and ($OutFile = "325 MXG")}
        ElseIf ($subfile4 -eq "6"){($directory = "$filepath"+"root\325 FW\325 OG") -and ($OutFile = "325 OG")}
        ElseIf ($subfile4 -eq "7"){($directory = "$filepath"+"root\325 FW\325 MSG\325 MSG Staff") -and ($OutFile = "MSG Staff")}
        ElseIf ($subfile4 -eq "8"){($directory = "$filepath"+"root\325 FW\325 MSG\325 MSG Public") -and ($OutFile = "MSG Public")}
        ElseIf ($subfile4 -eq "9"){($directory = "$filepath"+"root\325 FW\325 MSG\325 CES") -and ($OutFile = "325 CES")}
        ElseIf ($subfile4 -eq "10"){($directory = "$filepath"+"root\325 FW\325 MSG\325 CONS") -and ($OutFile = "325 CONS")}
        ElseIf ($subfile4 -eq "11"){($directory = "$filepath"+"root\325 FW\325 MSG\325 CS") -and ($OutFile = "325 CS")}
        ElseIf ($subfile4 -eq "12"){($directory = "$filepath"+"root\325 FW\325 MSG\325 FSS") -and ($OutFile = "325 FSS")}
        ElseIf ($subfile4 -eq "13"){($directory = "$filepath"+"root\325 FW\325 MSG\325 LRD") -and ($OutFile = "325 LRD")}
        ElseIf ($subfile4 -eq "14"){($directory = "$filepath"+"root\325 FW\325 MSG\325 SFS") -and ($OutFile = "325 SFS")}
        ElseIf ($subfile4 -eq "15"){($directory = "$filepath"+"root\325 FW\325 MXG\325 MXG Staff") -and ($OutFile = "MXG Staff")}
        ElseIf ($subfile4 -eq "16"){($directory = "$filepath"+"root\325 FW\325 MXG\325 MXG Public") -and ($OutFile = "MXG Public")}
        ElseIf ($subfile4 -eq "17"){($directory = "$filepath"+"root\325 FW\325 MXG\325 AMXS") -and ($OutFile = "325 AMXS")}
        ElseIf ($subfile4 -eq "18"){($directory = "$filepath"+"root\325 FW\325 MXG\325 MOS") -and ($OutFile = "325 MOS")}
        ElseIf ($subfile4 -eq "19"){($directory = "$filepath"+"root\325 FW\325 MXG\325 MXS") -and ($OutFile = "325 MXS")}
        ElseIf ($subfile4 -eq "20"){($directory = "$filepath"+"root\325 FW\325 MXG\372 TRS") -and ($OutFile = "372 TRS")}
        ElseIf ($subfile4 -eq "21"){($directory = "$filepath"+"root\325 FW\325 MXG\LOCMAR") -and ($OutFile = "LOCMAR")}
        ElseIf ($subfile4 -eq "22"){($directory = "$filepath"+"root\325 FW\325 OG\325 OG Staff") -and ($OutFile = "OG Staff")}
        ElseIf ($subfile4 -eq "23"){($directory = "$filepath"+"root\325 FW\325 OG\325 OG Public") -and ($OutFile = "OG Public")}
        ElseIf ($subfile4 -eq "24"){($directory = "$filepath"+"root\325 FW\325 OG\43 FS") -and ($OutFile = "43 FS")}
        ElseIf ($subfile4 -eq "25"){($directory = "$filepath"+"root\325 FW\325 OG\95 FS") -and ($OutFile = "95 FS")}
        ElseIf ($subfile4 -eq "26"){($directory = "$filepath"+"root\325 FW\325 OG\325 OSS") -and ($OutFile = "325 OSS")}
        ElseIf ($subfile4 -eq "27"){($directory = "$filepath"+"root\325 FW\325 OG\325 TRSS") -and ($OutFile = "325 TRSS")}
        ElseIf ($subfile4 -ne "0-27"){($directory = "$filepath"+"$subfile4") -and ($OutFile = "$subfile4")}
        }
Write-Host
$fileName = Read-Host "Filename or Filetype (ex. Test or .jpg)"
$Name = "$OutFile"+" ($fileName)"+" Files Found"+" $Date"
$body = "<TITLE>Files Found</TITLE><CENTER><H1>$Name</H1></CENTER>"
$a = "<style>"
$a = $a + "BODY{background-color:palegoldenrod;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:grey}"
$a = $a + "TD{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:white}"
$a = $a + "</style>"
Get-ChildItem -Recurse -Force $directory -ErrorAction SilentlyContinue |
Where-Object {($_.PSIsContainer -eq $false) -and  ($_.Name -like "*$fileName*")} | 
Select-Object Name,Directory | ConvertTo-HTML -head $body -body $a | Out-File "C:\Users\$Name.html"
$c = New-Object -Comobject wscript.shell
$b = $c.popup("The Search has Finished. All results are saved in the following location: C:\Users\$Name.html",0,"Complete",0)
Write-Host