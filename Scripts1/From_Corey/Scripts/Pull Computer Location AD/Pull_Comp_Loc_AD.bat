@ECHO OFF
@TITLE Pulling Computer Location from Active Directory

echo Computer Name: %computername%
echo Computer Name: %computername% >> C:\Temp\Quars2.txt

set locnow="dsget computer CN=%computername%,OU="Tyndall AFB Computers",OU="Tyndall AFB",OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL -loc'

set complocation=
for /F "delims=*" %%a in (%locnow%) do set complocation=%%a
echo %comploaction%
echo %complocation% >> C:\Temp\1111111.txt
pause