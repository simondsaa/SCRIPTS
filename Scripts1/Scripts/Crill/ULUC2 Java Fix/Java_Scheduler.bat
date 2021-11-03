@echo off
cls
schtasks /query > doh
findstr /B /I "ULUC2" doh >nul
if %errorlevel%==0  goto :delete
goto :create
 
:delete
SCHTASKS /DELETE /TN "ULUC2" /F >nul
 
:create
schtasks /create /TN "ULUC2" /SC ONSTART /TR "\\xlwu-fs-05pv\Tyndall_public\logons\uluc2\main.bat" /ru "System"

del doh >nul
