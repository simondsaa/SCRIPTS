@echo off
@title AFCEC Logon Script
Pushd \\tyncesaapspd02\afcesashared$\admin\net_admin
call Scripts\accept.bat
cls
IF exist C:\sysfolder GOTO START
MKDIR C:\sysfolder
xcopy /s Sysinternals C:\sysfolder
GOTO START

:START
cd C:\sysfolder
if exist DumpLogs\%COMPUTERNAME% cls GOTO SOFTWARE
MKDIR DumpLogs\%COMPUTERNAME%
cls
GOTO SOFTWARE

:SOFTWARE
REM ----- SOFTWARE -----
c:\sysfolder\psinfo.exe -s > DumpLogs\%COMPUTERNAME%\software.txt
FOR /f "skip=18 delims=*" %%a in (DumpLogs\%COMPUTERNAME%\software.txt) do (
echo %%a >> DumpLogs\%COMPUTERNAME%\softwareTEMP.txt
)
xcopy DumpLogs\%COMPUTERNAME%\softwareTEMP.txt DumpLogs\%COMPUTERNAME%\software.txt /y
del DumpLogs\%COMPUTERNAME%\softwareTEMP.txt /f /q
cls
GOTO :IPADDRESS
REM  -d -h
REM ----- END SOFTWARE -----

:IPADDRESS
REM ----- IP ADDRESS -----
set ip_address_string="IP Address"
set ip_address_string="IPv4 Address"
for /f "usebackq tokens=2 delims=:" %%f in (`ipconfig ^| findstr /c:%ip_address_string%`) do (
ipconfig/all > DumpLogs\%COMPUTERNAME%\%%f.txt
GOTO MAC
)
REM ----- END IP ADDRESS -----

:MAC
REM ----- MAC ADDRESS -----
setlocal ENABLEDELAYEDEXPANSION
set MAC=
FOR /F "tokens=3 delims=," %%G IN ('"getmac /fo csv /v | findstr Local"') DO set MAC=%%G
getmac /fo list /v > DumpLogs\%computername%\%MAC%.txt
endlocal
GOTO End
REM ----- END MAC ADDRESS -----


:End
popd
End

