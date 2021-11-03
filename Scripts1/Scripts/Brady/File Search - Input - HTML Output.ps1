#----------------------------------------------------------------------------------
#                           Written by SrA Timothy Brady
#                           Tyndall AFB, Panama City, FL
#                             Created January 28, 2014
#----------------------------------------------------------------------------------

# Please be sure to review the entire script before executing. Many changes will need ot be performed outside of Tyndall's enviroment.
# Also there is a line near the end that should be omited unless specifically needed for a specific search.

$FS1 = "\\XLWU-FS-001\"
$FS2 = "\\XLWU-FS-002\"
$FS3 = "\\XLWU-FS-003\"
$FS4 = "\\XLWU-FS-004\"

$AFNORTH = "$FS1"+"root\ANG"
$AFCEC = "$FS3"+"root\AFCESA"
$NCOA = "$FS2"+"root\NCOA"
$RHS = "$FS2"+"root\RHS"
$WEG = "$FS3"+"root\53WEG"
$337ACS = "$FS2"+"root\325 OG\325 ACS"
$372TRS = "$FS2"+"root\325 MXG\372 TRS"
$44FG = "$FS2"+"root\44 FG"

$325FW = "$FS4"+"root\325 FW"
$325MSG = "$325FW"+"\325 MSG"
$325MXG = "$325FW"+"\325 MXG"
$325OG = "$325FW"+"\325 OG"
$325CES = "$325MSG"+"\325 CES"
$325CONS = "$325MSG"+"\325 CONS"
$325CS = "$325MSG"+"\325 CS"
$325FSS = "$325MSG"+"\325 FSS"
$325LRD = "$325MSG"+"\325 LRD"
$325SFS = "$325MSG"+"\325 SFS"
$325AMXS = "$325MXG"+"\325 AMXS"
$325MXS = "$325MXG"+"\325 MXS"
$325LOCMAR = "$325MXG"+"\LOCMAR"
$325OSS = "$325OG"+"\325 OSS"
$325TRSS = "$325OG"+"\325TRSS"
$43FS = "$325OG"+"\43 FS"
$95FS = "$325OG"+"\95 FS"

Write-Host
Write-Host "0 = Server
1 = Tenant Unit
2 = Wing Unit"
Write-Host
$Unit = Read-Host "Would you like to search a Server, Tenant Unit, or Wing Unit"

If ($Unit -eq "0")
    {
    Write-Host
    Write-Host "Servers:
    
    1 = XLWU-FS-001
    2 = XLWU-FS-002
    3 = XLWU-FS-003
    4 = XLWU-FS-004"
    Write-Host
    $ServerSearch = Read-Host "Which Server would you like to search"
        
    If ($ServerSearch -eq "1"){($Directory = "$FS1"+"root" ) -and ($OutFile = "FS-001")}
    ElseIf ($ServerSearch -eq "2"){($Directory = "$FS2"+"root") -and ($OutFile = "FS-002")}
    ElseIf ($ServerSearch -eq "3"){($Directory = "$FS3"+"root") -and ($OutFile = "FS-003")}
    ElseIf ($ServerSearch -eq "4"){($Directory = "$FS4"+"root") -and ($OutFile = "FS-004")}
    ElseIf ($ServerSearch -ne "0-4"){Exit Write-Host "Incorrect selection, terminating."}
    }

ElseIf ($Unit -eq "1")
    {
    Write-Host
    Write-Host "Tenant Units:
    
    0 = AFNORTH
    1 = AFCEC
    2 = NCOA
    3 = 823 RHS
    4 = 53 WEG
    5 = 337 ACS
    6 = 372 TRS
    7 = 44 FG"
    Write-Host
    $TenantSearch = Read-Host "Which unit would you like to search"
                                                                
    If ($TenantSearch -eq "0"){($Directory = $AFNORTH) -and ($Outfile = "AFNORTH")}
    ElseIf ($TenantSearch -eq "1"){($Directory = $AFCEC) -and ($OutFile = "AFCEC")}
    ElseIf ($TenantSearch -eq "2"){($Directory = $NCOA) -and ($OutFile = "NCOA")}
    ElseIf ($TenantSearch -eq "3"){($Directory = $RHS) -and ($OutFile = "823 RHS")}
    ElseIf ($TenantSearch -eq "4"){($Directory = $WEG) -and ($OutFile = "53 WEG")}
    ElseIf ($TenantSearch -eq "5"){($Directory = $337ACS) -and ($OutFile = "337 ACS")}
    ElseIf ($TenantSearch -eq "6"){($Directory = $372TRS) -and ($OutFile = "372 TRS")}
    ElseIf ($TenantSearch -eq "7"){($Directory = $44FG) -and ($OutFile = "44 FG")}
    ElseIf ($TenantSearch -ne "0-7"){Exit Write-Host "Incorrect selection, terminating."}
    }

ElseIf ($Unit -eq "2")
    {
    Write-Host
    Write-Host "Wing Units:
    
    0 = 325 FW
    1 = 325 MSG
    2 = 325 MXG
    3 = 325 OG
    4 = 325 CES
    5 = 325 CONS
    6 = 325 CS
    7 = 325 FSS
    8 = 325 LRD
    9 = 325 SFS
    10 = 325 AMXS
    11 = 325 MXS
    12 = 325 LOCMAR
    13 = 325 OSS
    14 = 325 TRSS
    15 = 43 FS
    16 = 95 FS"
    Write-Host
    $WingSearch = Read-Host "Which unit would you like to search"

    If ($WingSearch -eq "0"){($Directory = $325FW) -and ($Outfile = "325 FW")}
    ElseIf ($WingSearch -eq "1"){($Directory = $325MSG) -and ($OutFile = "325 MSG")}
    ElseIf ($WingSearch -eq "2"){($Directory = $325MXG) -and ($OutFile = "325 MXG")}
    ElseIf ($WingSearch -eq "3"){($Directory = $325OG) -and ($OutFile = "325 OG")}
    ElseIf ($WingSearch -eq "4"){($Directory = $325CES) -and ($OutFile = "325 CES")}
    ElseIf ($WingSearch -eq "5"){($Directory = $325CONS) -and ($OutFile = "325 CONS")}
    ElseIf ($WingSearch -eq "6"){($Directory = $325CS) -and ($OutFile = "325 CS")}
    ElseIf ($WingSearch -eq "7"){($Directory = $325FSS) -and ($OutFile = "325 FSS")}
    ElseIf ($WingSearch -eq "8"){($Directory = $325LRD) -and ($OutFile = "325 LRD")}
    ElseIf ($WingSearch -eq "9"){($Directory = $325SFS) -and ($OutFile = "325 SFS")}
    ElseIf ($WingSearch -eq "10"){($Directory = $325AMXS) -and ($OutFile = "325 AMXS")}
    ElseIf ($WingSearch -eq "11"){($Directory = $325MXS) -and ($OutFile = "325 MXS")}
    ElseIf ($WingSearch -eq "12"){($Directory = $325LOCMAR) -and ($OutFile = "LOCMAR")}
    ElseIf ($WingSearch -eq "13"){($Directory = $325OSS) -and ($OutFile = "325 OSS")}
    ElseIf ($WingSearch -eq "14"){($Directory = $325TRSS) -and ($OutFile = "325 TRSS")}
    ElseIf ($WingSearch -eq "15"){($Directory = $43FS) -and ($OutFile = "43 FS")}
    ElseIf ($WingSearch -eq "16"){($Directory = $95FS) -and ($OutFile = "95 FS")}
    ElseIf ($WingSearch -ne "0-16"){Exit Write-Host "Incorrect selection, terminating."}
    }

Write-Host

$FileName = Read-Host "Filename or Filetype you want to search for ('Schedule' or '.jpg')"

$Date = Get-Date -UFormat "%d-%b-%g %H%M"
$Name = "$OutFile"+" ($FileName)"+" Files Found"+" $Date"

# HTML Formatting
$body = "<TITLE>Files Found</TITLE><CENTER><H1>$Name</H1></CENTER>"
$a = "<style>"
$a = $a + "BODY{background-color:palegoldenrod;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:lightgrey}"
$a = $a + "TD{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:white}"
$a = $a + "</style>"

$starttimer = Get-Date

# Search command begins
$Files = Get-ChildItem -Recurse -Force $Directory -ErrorAction SilentlyContinue | Where-Object {($_.PSIsContainer -eq $false) -and  ($_.Name -like "*$FileName*")} 

# Formats the results into the HTML Table
$Files | Select-Object Name,Directory | ConvertTo-HTML -head $body -body $a | Out-File "C:\$Name.html"

# Copies the found files into the below specified driectory. (Currently omitted) ***To include, remove # from the front of the next line***
#$Files | Copy-Item -Destination "C:\Users\timothy.brady\Desktop\Test"

$stoptimer = Get-Date
$ExecutionTime = [Math]::round(($stoptimer - $starttimer).TotalMinutes , 2)

$c = New-Object -Comobject wscript.shell
$b = $c.popup("The search has finished. Execution time: $ExecutionTime. All results are saved in the following location: C:\$Name.html",0,"Complete",80)
Write-Host