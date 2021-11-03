@ECHO OFF
SET ThisScriptsDirectory=C:\Users\1180219788A\Desktop
SET PowerShellScriptPath=C:\Users\1180219788A\Desktop\Kill_Advertisements.ps1
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%PowerShellScriptPath%""' -Verb RunAs}";