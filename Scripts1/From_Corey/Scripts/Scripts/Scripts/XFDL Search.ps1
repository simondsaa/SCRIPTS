Write-Host
Write-Host "----------------------------------------------------------------------------------
                           Written by SrA Timothy Brady
                           Tyndall AFB, Panama City, FL
----------------------------------------------------------------------------------"
$server = Read-Host "Which File Server Would You Like to Search? 1, 2, 3, or 4"
Write-Host
If ($server -eq "1"){$filepath = "\\XLWU-FS-001\root\"
Write-Host "Available Subfolders: 
 1 = 325 MSG\325 CES
 2 = 325 MSG\325 CS
 3 = 325 MSG\325 FSS
 4 = 325 MSG\325 LRD
 5 = 325 MSG\325 MSS
 6 = 325 MSG\325 SFS
 7 = 325 MSG\325 SVS
 8 = ANG\1AF
 9 = ANG\601 AOG
 10 = ANG\601 COD
 11 = ANG\601 CPD
 12 = ANG\601 SD
 13 = ANG\CONR
 14 = ANG\COS
 15 = ANG\IMO
 Or type the Path if known"
Write-Host
    $subfile1 = Read-Host "Subfolder *OPTIONAL*"
        If ($subfile1 -eq "1"){($directory = "$filepath"+"325 MSG\325 CES") -and ($OutFile = "325 CES")}
        ElseIf ($subfile1 -eq "2"){($directory = "$filepath"+"325 MSG\325 CS") -and ($OutFile = "325 CS")}
        ElseIf ($subfile1 -eq "3"){($directory = "$filepath"+"325 MSG\325 FSS") -and ($OutFile = "325 FSS")}
        ElseIf ($subfile1 -eq "4"){($directory = "$filepath"+"325 MSG\325 LRD") -and ($OutFile = "325 LRD")}
        ElseIf ($subfile1 -eq "5"){($directory = "$filepath"+"325 MSG\325 MSS") -and ($OutFile = "325 MSS")}
        ElseIf ($subfile1 -eq "6"){($directory = "$filepath"+"325 MSG\325 SFS") -and ($OutFile = "325 SFS")}
        ElseIf ($subfile1 -eq "7"){($directory = "$filepath"+"325 MSG\325 SVS") -and ($OutFile = "325 SVS")}
        ElseIf ($subfile1 -eq "8"){($directory = "$filepath"+"ANG\1AF") -and ($OutFile = "1 AF")}
        ElseIf ($subfile1 -eq "9"){($directory = "$filepath"+"ANG\601 AOG") -and ($OutFile = "601 AOG")}
        ElseIf ($subfile1 -eq "10"){($directory = "$filepath"+"ANG\601 COD") -and ($OutFile = "601 COD")}
        ElseIf ($subfile1 -eq "11"){($directory = "$filepath"+"ANG\601 CPD") -and ($OutFile = "601 CPD")}
        ElseIf ($subfile1 -eq "12"){($directory = "$filepath"+"ANG\601 SD") -and ($OutFile = "601 SD")}
        ElseIf ($subfile1 -eq "13"){($directory = "$filepath"+"ANG\CONR") -and ($OutFile = "CONR")}
        ElseIf ($subfile1 -eq "14"){($directory = "$filepath"+"ANG\COS") -and ($OutFile = "COS")}
        ElseIf ($subfile1 -eq "15"){($directory = "$filepath"+"ANG\IMO") -and ($OutFile = "IMO")}
        ElseIf ($subfile1 -ne "1-15"){($directory = "$filepath"+"$subfile1") -and ($OutFile = "$subfile1")}
        }
ElseIf ($server -eq "2"){$filepath = "\\XLWU-FS-002\root\"
Write-Host "Available Subfolders:
 1 = 325 FW\Shared
 2 = 325 MXG\325 AMXS
 3 = 325 MXG\325 MOS
 4 = 325 MXG\325 MXS
 5 = 325 MXG\372 TRS
 6 = 325 OG\43 FS
 7 = 325 OG\95 FS
 8 = 325 OG\325 ACS
 9 = 325 OG\325 OSS
 10 = 325 OG\325 TRSS
 11 = NCOA\Shared
 12 = RHS\Shared
 13 = tyndall
 Or type the Path if known"
Write-Host
    $subfile2 = Read-Host "Subfolder *OPTIONAL*"
        If ($subfile2 -eq "1"){($directory = "$filepath"+"325 FW\Shared") -and ($OutFile = "325 FW")}
        ElseIf ($subfile2 -eq "2"){($directory = "$filepath"+"325 MXG\325 AMXS") -and ($OutFile = "325 AMXS")}
        ElseIf ($subfile2 -eq "3"){($directory = "$filepath"+"325 MXG\325 MOS") -and ($OutFile = "325 MOS")}
        ElseIf ($subfile2 -eq "4"){($directory = "$filepath"+"325 MXG\325 MXS") -and ($OutFile = "325 MXS")}
        ElseIf ($subfile2 -eq "5"){($directory = "$filepath"+"325 MXG\372 TRS") -and ($OutFile = "372 TRS")}
        ElseIf ($subfile2 -eq "6"){($directory = "$filepath"+"325 OG\43 FS") -and ($OutFile = "43 FS")}
        ElseIf ($subfile2 -eq "7"){($directory = "$filepath"+"325 OG\95 FS") -and ($OutFile = "95 FS")}
        ElseIf ($subfile2 -eq "8"){($directory = "$filepath"+"325 OG\325 ACS") -and ($OutFile = "325 ACS")}
        ElseIf ($subfile2 -eq "9"){($directory = "$filepath"+"325 OG\325 OSS") -and ($OutFile = "325 OSS")}
        ElseIf ($subfile2 -eq "10"){($directory = "$filepath"+"325 OG\325 TRSS") -and ($OutFile = "325 TRSS")}
        ElseIf ($subfile2 -eq "11"){($directory = "$filepath"+"NCOA\Shared") -and ($OutFile = "NCOA")}
        ElseIf ($subfile2 -eq "12"){($directory = "$filepath"+"RHS\Shared") -and ($OutFile = "823 RHS")}
        ElseIf ($subfile2 -eq "13"){($directory = "$filepath"+"tyndall") -and ($OutFile = "tyndall")}
        ElseIf ($subfile2 -ne "1-13"){($directory = "$filepath"+"$subfile2") -and ($OutFile = "$subfile2")}
        }
ElseIf ($server -eq "3"){$filepath = "\\XLWU-FS-003\root\"
Write-Host "Available Subfolders:
 1 = 53WEG\Shared
 2 = 361 TRS\361 TRS
 3 = 479_FTG\451FTS
 4 = 479_FTG\455FTS
 5 = 479_FTG\479FTG
 6 = 479_FTG\479OSS
 7 = AFCESA\Shared
 8 = Cons\Shared
 9 = tyndall
 Or type the Path if known"
Write-Host
    $subfile3 = Read-Host "Subfolder *OPTIONAL*"
        If ($subfile3 -eq "1"){($directory = "$filepath"+"53WEG\Shared") -and ($OutFile = "53 WEG")}
        ElseIf ($subfile3 -eq "2"){($directory = "$filepath"+"361 TRS\361 TRS") -and ($OutFile = "361 TRS")}
        ElseIf ($subfile3 -eq "3"){($directory = "$filepath"+"479_FTG\451FTS") -and ($OutFile = "451 FTS")}
        ElseIf ($subfile3 -eq "4"){($directory = "$filepath"+"479_FTG\455FTS") -and ($OutFile = "455 FTS")}
        ElseIf ($subfile3 -eq "5"){($directory = "$filepath"+"479_FTG\479FTG") -and ($OutFile = "479 FTG")}
        ElseIf ($subfile3 -eq "6"){($directory = "$filepath"+"479_FTG\479OSS") -and ($OutFile = "479 OSS")}
        ElseIf ($subfile3 -eq "7"){($directory = "$filepath"+"AFCESA\Shared") -and ($OutFile = "AFCESA")}
        ElseIf ($subfile3 -eq "8"){($directory = "$filepath"+"Cons\Shared") -and ($OutFile = "Cons")}
        ElseIf ($subfile3 -eq "9"){($directory = "$filepath"+"tyndall") -and ($OutFile = "tyndall")}
        ElseIf ($subfile3 -ne "1-9"){($directory = "$filepath"+"$subfile3") -and ($OutFile = "$subfile3")}
        }
ElseIf ($server -eq "4"){$filepath = "\\XLWU-FS-004\root\"
Write-Host "Available Subfolders:
 1 = LMShared\LMIS
 Or type the Path if known"
Write-Host
    $subfile4 = Read-Host "Subfolder *OPTIONAL*"
        If ($subfile4 -eq "1"){($directory = "$filepath"+"LMShared\LMIS") -and ($OutFile = "LMIS")}
        ElseIf ($subfile4 -ne "1"){($directory = "$filepath"+"$subfile4") -and ($OutFile = "$subfile4")}
        }
Write-Host
$Name = "$OutFile"+" XFDL Files Found"
$body = "<TITLE>Files Found</TITLE><CENTER><H1>$Name</H1></CENTER>"
$a = "<style>"
$a = $a + "BODY{background-color:palegoldenrod;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:grey}"
$a = $a + "TD{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:white}"
$a = $a + "</style>"
$xfdl_docs = Get-ChildItem -Recurse -Force $directory -include "*.xfdl" -ErrorAction SilentlyContinue
$XfdlSearchStrings = @("[EO]PR","707","910","911","55","4392")
ForEach ($Xfdl in $xfdl_docs){
    $epr	= $Xfdl -match $XfdlSearchStrings[0]
    $f707	= $Xfdl -match $XfdlSearchStrings[1]
    $f910	= $Xfdl -match $XfdlSearchStrings[2]
	$f911	= $Xfdl -match $XfdlSearchStrings[3]
	$f55	= $Xfdl -match $XfdlSearchStrings[4]
    $f4392  = $xfdl -match $XfdlSearchStrings[5]}
If(($epr) -or ($f707) -or ($f910) -or ($f911) -or ($f55) -or ($f4392)){
$Xfdl | Select Name,Directory | ConvertTo-HTML -head $body -body $a | Out-File "C:\Users\$Name.html" -append
}
Write-Host