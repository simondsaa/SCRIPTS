@echo off
@title AFCEC Logon Script

REM *******************************************
REM *     Last Modified By: Corey Jarrett     *
REM *      Last Modified On: 22-FEB-2013      *
REM *******************************************

set netadmin=\\tyncesaapspd02\afcesashared$\admin\net_admin
Pushd %netadmin%
call %netadmin%\Scripts\accept.bat
cls
IF exist C:\sysfolder GOTO START
MKDIR C:\sysfolder
xcopy /s %netadmin%\Sysinternals C:\sysfolder
GOTO START

REM *******************************************
REM *           Create Directories            *
REM *******************************************

:START
cd C:\sysfolder
if exist %netadmin%\DumpLogs\Computers\ByComputerName\%COMPUTERNAME% cls GOTO IFSERIAL
MKDIR %netadmin%\DumpLogs\Computers\ByComputerName\%COMPUTERNAME%
cls
GOTO IFSERIAL

:IFSERIAL

for /F "skip=2 tokens=2 delims=," %%A in ('wmic systemenclosure get serialnumber /FORMAT:csv') do (set "GetSerial=%%A")
GOTO SERIALGET

:SERIALGET
if exist %netadmin%\DumpLogs\Computers\BySerial\%GetSerial% cls GOTO IFUSER
MKDIR %netadmin%\DumpLogs\Computers\BySerial\%GetSerial%
cls
GOTO IFUSER

:IFUSER
if exist %netadmin%\DumpLogs\Computers\ByUser\%USERNAME% cls GOTO SOFTWARE
MKDIR %netadmin%\DumpLogs\Computers\ByUser\%USERNAME%
cls
GOTO SOFTWARE

:SOFTWARE
REM ----- SOFTWARE -----
c:\sysfolder\psinfo.exe -s > %netadmin%\DumpLogs\Computers\ByComputerName\%COMPUTERNAME%\software-%COMPUTERNAME%.txt
FOR /f "skip=18 delims=*" %%a in (%netadmin%\DumpLogs\Computers\ByComputerName\%COMPUTERNAME%\software-%COMPUTERNAME%.txt) do (
echo %%a >> %netadmin%\DumpLogs\Computers\software%COMPUTERNAME%TEMP.txt
)
echo f | xcopy %netadmin%\DumpLogs\Computers\software%COMPUTERNAME%TEMP.txt %netadmin%\DumpLogs\Computers\ByComputerName\%COMPUTERNAME%\software-%COMPUTERNAME%.txt /y
echo f | xcopy %netadmin%\DumpLogs\Computers\software%COMPUTERNAME%TEMP.txt "%netadmin%\DumpLogs\Computers\BySerial\%GetSerial%\software-%COMPUTERNAME%.txt" /y
echo f | xcopy %netadmin%\DumpLogs\Computers\software%COMPUTERNAME%TEMP.txt %netadmin%\DumpLogs\Computers\ByUser\%USERNAME%\software-%COMPUTERNAME%.txt /y
del %netadmin%\DumpLogs\Computers\software%COMPUTERNAME%TEMP.txt /f /q
cls
REM ----- END SOFTWARE -----

GOTO IPADDRESS

:IPADDRESS
REM ----- IP ADDRESS -----
set ip_address_string="IP Address"
set ip_address_string="IPv4 Address"
for /f "usebackq tokens=2 delims=:" %%f in (`ipconfig ^| findstr /c:%ip_address_string%`) do (
ipconfig/all > %netadmin%\DumpLogs\Computers\ByComputerName\%COMPUTERNAME%\%%f.txt
GOTO USERNAME
)
REM ----- END IP ADDRESS -----

:USERNAME
REM ----- USER NAME -----
echo Logged on %date% @ %time% >> %netadmin%\DumpLogs\Computers\ByComputerName\%COMPUTERNAME%\%username%.txt
GOTO MAC
REM ----- END USER NAME -----

:MAC
REM ----- MAC ADDRESS -----
setlocal ENABLEDELAYEDEXPANSION
set MAC=
FOR /F "tokens=3 delims=," %%G IN ('"getmac /fo csv /v | findstr Gigabit"') DO set MAC=%%G
getmac /fo list /v > %netadmin%\DumpLogs\Computers\ByComputerName\%COMPUTERNAME%\%MAC%.txt
endlocal
REM ----- END MAC ADDRESS -----

:NEWSERIAL
call %netadmin%\Scripts\NewCompName.vbs
set content=
for /f "delims=" %%i in ('type %netadmin%\DumpLogs\Computers\ByComputerName\%COMPUTERNAME%\tempCompName.txt') do set content=%%i
for /F "skip=2 tokens=2 delims=," %%A in ('wmic systemenclosure get serialnumber /FORMAT:csv') do (set "serial=%%A")
set serial2=52XLWU%content%3-%serial:~-7%
REM set realserial=52XLWU%content%3-%serial2%
echo %USERNAME% logged on to %COMPUTERNAME%/%serial2% on %date% @ %time% >> %netadmin%\DumpLogs\Computers\ByComputerName\%COMPUTERNAME%\%serial2%.txt
echo %USERNAME% logged on to %COMPUTERNAME%/%serial2% on %date% @ %time% >> %netadmin%\DumpLogs\Computers\BySerial\%GetSerial%\%serial2%.txt
echo %USERNAME% logged on to %COMPUTERNAME%/%serial2% on %date% @ %time% >> %netadmin%\DumpLogs\Computers\ByUser\%USERNAME%\%serial2%.txt

del %netadmin%\DumpLogs\Computers\ByComputerName\%COMPUTERNAME%\tempCompName.txt

REM GOTO STARTPRINT

GOTO MAPDRIVES

:MAPDRIVES
REM echo y | net use L: /d
net use L: "\\tyncesaapspd02\AFCESA$" /PERSISTENT:YES
cls
REM echo y | net use R: /d
net use R: "\\tyncesaapspd02\afcesashared$\REFERENCE" /PERSISTENT:YES
cls
REM echo y | net use S: /d
net use S: "\\tyncesaapspd02\afcesashared$\SHARE" /PERSISTENT:YES
cls
REM echo y | net use T: /d
net use T: "\\xlwu-fs-002\TYNDALL$" /PERSISTENT:YES
cls

REM *******************************************
REM *          Custom Drive Mappings          *
REM *            Corey 26-Feb-2013            *
REM *******************************************

REM DELETE THIS STUFF LATER //Corey 26-Feb-2013
if %username%==michael.cowan GOTO P_CO_MAP

if %username%==corey.jarrett GOTO COREYMAP
if %username%==George.Vansteenburg GOTO VANSTEEN
if %username%==James.Garred GOTO V_MAP
if %username%==ryan.warnock.ctr GOTO MAPWARNOCK
if %username%==michael.clawson2 GOTO COTEMPMAP
if %username%==wes.somers GOTO COTEMPMAPSOMERS
if %username%==judy.biddle GOTO COTEMPMAP
if %username%==tracy.coughlin GOTO COTEMPMAP
if %username%==michael.cowan GOTO COTEMPMAP
if %username%==David.Dunne GOTO CNTEMPMAP
if %username%==michael.giniger GOTO CNTEMPMAP
if %username%==bil.hawkins GOTO COTEMPMAP
if %username%==Sundae.Knight GOTO CNTEMPMAP
if %username%==kari.kubista GOTO CNTEMPMAP
if %username%==Jason.Poe GOTO CNTEMPMAP
if %username%==joe.worrell GOTO COTEMPMAP
if %username%==richard.rude GOTO AFNORTH_AUGS
if %username%==steven.reed GOTO AFNORTH_AUGS
if %username%==marita.woods GOTO AFNORTH_AUGS
if %username%==samuel.hazzard GOTO AFNORTH_AUGS
if %username%==kimberly.bottomy GOTO BOTTOMMAP
if %username%==Robert.Mellerski GOTO MELLMAP
if %username%==Cherry.Roberts GOTO CHERRYMAP


GOTO CHECKDRIVES

:STARTPRINT
FIND /I "%username%" %netadmin%\Scripts\AlreadyInstalled.txt
if errorlevel 1 ( GOTO SETPRINT
) else (
GOTO End )



:SETPRINT
cls
FIND /I "%username%" %netadmin%\Scripts\PrinterInstallDump.csv > %netadmin%\Scripts\%username%_1.txt

set lines=2
set curr=1
for /f "tokens=1,2,3 delims=," %%a in (%netadmin%\Scripts\%username%_1.txt) do (
     for %%a in (!lines!) do (
         set cubicle=%%b&set print=%%c
)
set /a "curr = curr + 1"
)
for %%a in (%print%) do (
%windir%\system32\cscript.exe "%netadmin%\Scripts\printinstall.vbs" /prnr:"%print%"
GOTO Addname
)
GOTO End

:Addname
echo %username% - %date% >> %netadmin%\Scripts\AlreadyInstalled.txt
GOTO DELFILE
)
GOTO DELFILE
:DELFILE
del %netadmin%\Scripts\%username%_1.txt
GOTO END

REM *******************************************
REM *      Custom Drive Mappings (cont.)      *
REM *            Corey 26-Feb-2013            *
REM *******************************************

:CHERRYMAP
net use V: "\\tyncesaapspd02\afcesashared$\cebc" /PERSISTENT:YES
cls
GOTO CHECKDRIVES

:BOTTOMMAP
net use P: /d
net use P: "\\tyncesaapspd02\afcesashared$\CEK" /PERSISTENT:YES
cls
GOTO CHECKDRIVES

:MELLMAP
net use O: /d
net use O: "\\TYNAFRLAP60502\AFRLCommon$" /PERSISTENT:YES
cls
GOTO CHECKDRIVES

:COREYMAP
net use A: "\\xlwu-fs-003\AFCESAShared" /PERSISTENT:YES
cls
net use I: "\\xlwu-fs-003\AFCESAShared\Home\corey.jarrett" /PERSISTENT:YES
cls
net use P: "\\tyncesaapspd02\afcesa$" /PERSISTENT:YES
cls
net use P: "\\tyncesaapspd02\afcesashared$\CEB" /PERSISTENT:YES
cls
GOTO CHECKDRIVES

:VANSTEEN
net use P: "\\tyncesaapspd02\afcesashared$\CEO" /PERSISTENT:YES
cls
net use M: "\\tyncesaapspd02\afcesashared$\CEOP" /PERSISTENT:YES
cls
net use O: "\\tyncesaapspd02\afcesashared$\APE" /PERSISTENT:YES
cls
net use Q: "\\tyncesaapspd02\afcesashared$\CEX" /PERSISTENT:YES
cls
GOTO CHECKDRIVES

:MAPWARNOCK
net use P: "\\tyncesaapspd02\afcesashared$\A7CRT" /PERSISTENT:YES
cls
GOTO CHECKDRIVES

:V_MAP
net use V: "\\tyncesaapspd02\afcesashared$\CEBC" /PERSISTENT:YES
cls
GOTO CHECKDRIVES

:P_CO_MAP
net use P: "\\tyncesaapspd02\afcesashared$\CEO" /PERSISTENT:YES
cls
net use Q: "\\tyncesaapspd02\afcesashared$\CEN" /PERSISTENT:YES
cls
GOTO CHECKDRIVES

:COTEMPMAP
net use P: "\\tyncesaapspd02\afcesashared$\CEO" /PERSISTENT:YES
cls
net use Q: "\\tyncesaapspd02\afcesashared$\CEN" /PERSISTENT:YES
cls
GOTO CHECKDRIVES

:COTEMPMAPSOMERS
net use M: "\\tyncesaapspd02\afcesashared$\CEN" /PERSISTENT:YES
cls
GOTO CHECKDRIVES

:CNTEMPMAP
net use P: "\\tyncesaapspd02\afcesashared$\CEN" /PERSISTENT:YES
cls
net use Q: "\\tyncesaapspd02\afcesashared$\CEO" /PERSISTENT:YES
cls
GOTO CHECKDRIVES

:AFNORTH_AUGS
net USE H: "\\XLWU-FS-001\ANG$\Shared\_03 AOC\ACOMS\How to Videos" /PERSISTENT:YES
cls
net USE K: "\\XLWU-FS-001\ANG$\Shared" /PERSISTENT:YES
cls
net USE M: "\\XLWU-FS-001\AFNORTH_Media" /PERSISTENT:YES
cls
net USE J: "\\XLWU-FS-001\pfps$" /PERSISTENT:YES
cls
net use P: "\\tyncesaapspd02\afcesashared$\CEX" /PERSISTENT:YES
cls
GOTO CHECKDRIVES

GOTO CHECKDRIVES

:CHECKDRIVES
net use > %netadmin%\DumpLogs\Computers\ByComputerName\%COMPUTERNAME%\Drives-%USERNAME%.txt
net use > %netadmin%\DumpLogs\Computers\BySerial\%GetSerial%\Drives-%USERNAME%.txt
net use > %netadmin%\DumpLogs\Computers\ByUser\%USERNAME%\Drives-%COMPUTERNAME%.txt

GOTO END
:End
popd
End

