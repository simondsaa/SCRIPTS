@echo off
@title AFCEC Logon Script

if %username%==corey.jarrett GOTO TESTINGS
if %username%==rik.hix GOTO TESTINGS
if %username%==Christopher.Malott GOTO TESTINGS
if %username%==david.farmer GOTO TESTINGS

:KICKOFF
Pushd \\tyncesaapspd02\afcesashared$\admin\net_admin
call \\tyncesaapspd02\afcesashared$\admin\net_admin\Scripts\accept.bat
cls
IF exist C:\sysfolder GOTO START
MKDIR C:\sysfolder
xcopy /s \\tyncesaapspd02\afcesashared$\admin\net_admin\Sysinternals C:\sysfolder
GOTO START

:START
cd C:\sysfolder
if exist \\tyncesaapspd02\afcesashared$\admin\net_admin\DumpLogs\%COMPUTERNAME% cls GOTO SOFTWARE
MKDIR \\tyncesaapspd02\afcesashared$\admin\net_admin\DumpLogs\%COMPUTERNAME%
cls
GOTO SOFTWARE

:SOFTWARE
REM ----- SOFTWARE -----
c:\sysfolder\psinfo.exe -s > \\tyncesaapspd02\afcesashared$\admin\net_admin\DumpLogs\%COMPUTERNAME%\software.txt
FOR /f "skip=18 delims=*" %%a in (\\tyncesaapspd02\afcesashared$\admin\net_admin\DumpLogs\%COMPUTERNAME%\software.txt) do (
echo %%a >> \\tyncesaapspd02\afcesashared$\admin\net_admin\DumpLogs\%COMPUTERNAME%\softwareTEMP.txt
)
xcopy \\tyncesaapspd02\afcesashared$\admin\net_admin\DumpLogs\%COMPUTERNAME%\softwareTEMP.txt \\tyncesaapspd02\afcesashared$\admin\net_admin\DumpLogs\%COMPUTERNAME%\software.txt /y
del \\tyncesaapspd02\afcesashared$\admin\net_admin\DumpLogs\%COMPUTERNAME%\softwareTEMP.txt /f /q
cls
GOTO :IPADDRESS
REM ----- END SOFTWARE -----

:IPADDRESS
REM ----- IP ADDRESS -----
set ip_address_string="IP Address"
set ip_address_string="IPv4 Address"
for /f "usebackq tokens=2 delims=:" %%f in (`ipconfig ^| findstr /c:%ip_address_string%`) do (
ipconfig/all > \\tyncesaapspd02\afcesashared$\admin\net_admin\DumpLogs\%COMPUTERNAME%\%%f.txt
GOTO USERNAME
)
REM ----- END IP ADDRESS -----

:USERNAME
REM ----- USER NAME -----
echo Logged on %date% @ %time% >> \\tyncesaapspd02\afcesashared$\admin\net_admin\DumpLogs\%computername%\%username%.txt
GOTO MAC
REM ----- END USER NAME -----

:MAC
REM ----- MAC ADDRESS -----
setlocal ENABLEDELAYEDEXPANSION
set MAC=
FOR /F "tokens=3 delims=," %%G IN ('"getmac /fo csv /v | findstr Gigabit"') DO set MAC=%%G
getmac /fo list /v > \\tyncesaapspd02\afcesashared$\admin\net_admin\DumpLogs\%computername%\%MAC%.txt
endlocal

REM GOTO STARTPRINT

GOTO MAPDRIVES

:MAPDRIVES
net use L: "\\tyncesafsedie03\AFCESA$" /PERSISTENT:YES
net use R: "\\tyncesaapspd02\afcesashared$\REFERENCE" /PERSISTENT:YES
net use S: "\\tyncesaapspd02\afcesashared$\SHARE" /PERSISTENT:YES
net use T: "\\xlwu-fs-002\TYNDALL$" /PERSISTENT:YES

GOTO END
REM ----- END MAC ADDRESS -----

:STARTPRINT
FIND /I "%username%" \\tyncesaapspd02\afcesashared$\admin\net_admin\Scripts\AlreadyInstalled.txt
if errorlevel 1 ( GOTO SETPRINT
) else (
GOTO End )



:SETPRINT
cls
FIND /I "%username%" \\tyncesaapspd02\afcesashared$\admin\net_admin\Scripts\PrinterInstallDump.csv > \\tyncesaapspd02\afcesashared$\admin\net_admin\Scripts\%username%_1.txt

set lines=2
set curr=1
for /f "tokens=1,2,3 delims=," %%a in (\\tyncesaapspd02\afcesashared$\admin\net_admin\Scripts\%username%_1.txt) do (
     for %%a in (!lines!) do (
         set cubicle=%%b&set print=%%c
)
set /a "curr = curr + 1"
)
for %%a in (%print%) do (
%windir%\system32\cscript.exe "\\tyncesaapspd02\afcesashared$\admin\net_admin\Scripts\printinstall.vbs" /prnr:"%print%"
GOTO Addname
)
GOTO End

:Addname
echo %username% - %date% >> \\tyncesaapspd02\afcesashared$\admin\net_admin\Scripts\AlreadyInstalled.txt
GOTO DELFILE
)
GOTO DELFILE
:DELFILE
del \\tyncesaapspd02\afcesashared$\admin\net_admin\Scripts\%username%_1.txt
GOTO END

:TESTINGS
call \\tyncesaapspd02\afcesashared$\admin\net_admin\AFCECLogonTEST02012013.bat
GOTO EOF
:End
popd
End

