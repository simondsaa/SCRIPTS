@echo off
@title New AFNET Computer Name
mkdir c:\tempDirectory
call NewCompName.vbs
set content=
for /f "delims=" %%i in ('type c:\tempDirectory\tempCompName.txt') do set content=%%i
for /F "skip=2 tokens=2 delims=," %%A in ('wmic systemenclosure get serialnumber /FORMAT:csv') do (set "serial=%%A")
set serial2=%serial:~-7%
echo 52XLWU%content%3-%serial2% > c:\tempDirectory\ComputerName_PleaseClose.txt
call c:\tempDirectory\ComputerName_PleaseClose.txt
rmdir /s /q c:\tempDirectory