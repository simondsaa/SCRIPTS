@ECHO OFF
@TITLE AFNET VPN LOGON SCRIPT ADDITION
PUSHD

REM AFNET VPN LOGON SCRIPT ADDITION by C. Jarrett, 101 ACOMS CS, Tyndall AFB, DSN 742-0272

REM *** Copies Welcome.vbs To Primary Directory ***
xcopy "\\xlwu-fs-001\ANG$\Shared\_03 AOC\ACOMS\SCO\SCOC\CSA\SOFTWARE\VPN\INSTALL_AFNET_VPN\welcome.vbs" "C:\ProgramData\Microsoft\Network\Connections\Cm\L2TP32\" /Y

REM *** Make Working Directory ***
mkdir c:\VPN_TEMP_FOLDER

REM *** Change Directory For Next Command ***
cd c:\users

REM *** Output User List To Temp File ***
dir /b > C:\VPN_TEMP_FOLDER\TEMP.txt

REM *** Loop Through List To Delete All Instances Of Welcome.vbs ***
for /F "tokens=*" %%A in (C:\VPN_TEMP_FOLDER\TEMP.txt) do (IF EXIST "C:\Users\%%A\AppData\Roaming\Microsoft\Network\Connections\_hiddencm\" rmdir /q /s "C:\Users\%%A\AppData\Roaming\Microsoft\Network\Connections\_hiddencm")

REM *** Delete Temp Working Directory ***
rmdir /q /s C:\VPN_TEMP_FOLDER

POPD
REM *** End Of File ***