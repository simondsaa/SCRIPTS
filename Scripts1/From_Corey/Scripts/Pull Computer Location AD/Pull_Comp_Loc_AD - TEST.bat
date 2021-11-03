{
echo Computer Name: %computername%
echo Computer Name: %computername% >> C:\Users\1252862141.adm\Desktop\Scripts\Active_Directory_Pull.txt
}
{
set locnow="dsget computer CN=%computername%,OU="Tyndall AFB Computers",OU="Tyndall AFB",OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL -loc'

set complocation=
for /F "delims=*" %%a in (%locnow%) do set complocation=%%a
echo %comploaction%
echo %complocation% >> C:\Users\1252862141.adm\Desktop\Scripts\Active_Directory_Pull.txt
pause"
}
