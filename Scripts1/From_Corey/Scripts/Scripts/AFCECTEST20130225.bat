@echo off
@title AFCEC Logon Script

REM Last Modified By: Corey Jarrett
REM Last Modified On: 22-FEB-2013

set netadmin=\\tyncesaapspd02\afcesashared$\admin\net_admin
Pushd %netadmin%
call %netadmin%\Scripts\accept.bat
cls
IF exist C:\sysfolder GOTO START
MKDIR C:\sysfolder
xcopy /s %netadmin%\Sysinternals C:\sysfolder
GOTO START

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
net use L: "\\tyncesafsedie03\AFCESA$" /PERSISTENT:YES
cls
net use R: "\\tyncesaapspd02\afcesashared$\REFERENCE" /PERSISTENT:YES
cls
net use S: "\\tyncesaapspd02\afcesashared$\SHARE" /PERSISTENT:YES
cls
net use T: "\\xlwu-fs-002\TYNDALL$" /PERSISTENT:YES
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
GOTO END

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
:End
popd
End

