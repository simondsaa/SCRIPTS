Echo "Calling EMWProf.exe" %time% >> "%userprofile%\emwprofbat.log"

"%LOGONSERVER%\NETLOGON\EMWProf\EMWProf.exe" -Log "\\52XLWU-MG-001v\EMWProfLogs\%username%_#h_#d_#t.log" -ini "\\52XLWU-mg-001v\EMWProfLogs\INI\EMWProf.ini"

Echo "Returning to the EMWProf.vbs script" %time% >> "%userprofile%\emwprofbat.log"