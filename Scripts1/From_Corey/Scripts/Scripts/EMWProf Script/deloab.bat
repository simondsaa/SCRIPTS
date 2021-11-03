echo off

REM OS Check

ver | find "5.1.2600" > nul
if %errorlevel% == 0 goto XP

ver | find "5.2.3790" > nul
if %errorlevel% == 0 goto XP

ver | find "6.0.6002" > nul
if %errorlevel% == 0 goto Vista/7

ver | find "6.0.6000" > nul
if %errorlevel% == 0 goto Vista/7

ver | find "6.1.7600" > nul
if %errorlevel% == 0 goto Vista/7

ver | find "6.1.7601" > nul
if %errorlevel% == 0 goto Vista/7

REM Remove OST/TMP Files

:XP
cd "%userprofile%\Local Settings\Application Data\Microsoft\Outlook"
del *.oab /s /f /q
del *.tmp /f /q
goto end

:Vista/7
cd %userprofile%\AppData\Local\Microsoft\Outlook
del *.oab /s /f /q
del *.tmp /f /q
goto end

REM End
:end
