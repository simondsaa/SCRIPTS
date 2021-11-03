REM @echo off


:: Use date /t and time /t from the command line to get the format of your date and
:: time; change the substring below as needed.

:: This will create a timestamp like yyyy-mm-dd-hh-mm-ss.
set TIMESTAMP=%DATE:~10,4%-%DATE:~4,2%-%DATE:~7,2%-%TIME:~0,2%-%TIME:~3,2%-%TIME:~6,2%

@echo TIMESTAMP=%TIMESTAMP%

:: Create a new directory
md "C:\temp\%TIMESTAMP%"


set filelocation=c:\temp\jacce.txt
set fileoutput=C:\temp\%TIMESTAMP%


echo Starting Software List Pull
FOR /F "tokens=*" %%A in (%filelocation%) do (wmic /node:"%%A" product get name, version /format:csv > %fileoutput%\SL_%%A.csv)
pause