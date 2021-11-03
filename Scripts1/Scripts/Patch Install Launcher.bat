@ECHO OFF
SET ThisScriptsDirectory=C:\Users\1180219788A\Downloads\Scripts
SET PowerShellScriptPath=C:\Users\1180219788A\Downloads\Scripts\Job.ps1
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%PowerShellScriptPath%""' -Verb RunAs}";