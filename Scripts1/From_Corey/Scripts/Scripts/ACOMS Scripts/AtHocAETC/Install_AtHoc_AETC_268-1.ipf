Document Type: IPF
item: Global
  Version=6.0
  Title English=Install AtHoc AETC 268
  Flags=00000100
  Languages=0 0 65 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
  LanguagesList=English
  Japanese Font Name=MS Gothic
  Japanese Font Size=10
  Start Gradient=0 0 255
  End Gradient=0 0 0
  Windows Flags=00010100000000010010110000011000
  Log Pathname=%WIN%\temp\AtHocAETC.log
  Message Font=MS Sans Serif
  Font Size=8
  Disk Filename=SETUP
  Patch Flags=0000000000000001
  Patch Threshold=85
  Patch Memory=4000
  File Version=2.0.148.00
  File Description=Microsoft Systems Management Server Installer
  Copyright=Copyright (C) Microsoft Corp. 1997-2001
  Company Name=Microsoft Corporation
  Internal Name=Smsistub
  Original FileName=Stub.exe
  Product Name=Microsoft Systems Management Server Installer
  Product Version=2.0.148.00
  FTP Cluster Size=20
end
item: Open/Close INSTALL.LOG
  Pathname=c:\Windows\temp\Install_atHoc_268.log
  Flags=00000010
end
item: Set Variable
  Variable=SYS
  Value=%WIN%
end
item: Open/Close INSTALL.LOG
end
item: Set Variable
  Variable=ATHOCLOG
  Value=/l*vx c:\Windows\temp\AtHocAETC6_2_11_268.log
end
item: Set Variable
  Variable=BASEURL
  Value=BASEURL=https://alerts.aetc.af.mil/config/baseurl.asp
end
item: Set Variable
  Variable=SUFFIX
  Value=RUNAFTERINSTALL=N DESKBAR=N TOOLBAR=N SILENT=Y VALIDATECERT=y MANDATESSL=y UNINSTALLOPTION=N
end
item: Set Variable
  Variable=APPTITLE
  Value=Install AtHoc
  Flags=10000000
end
item: Set Variable
  Variable=GROUP
  Flags=10000000
end
item: Set Variable
  Variable=DISABLED
  Value=!
end
item: Parse String
  Source=%WIN%
  Pattern=:
  Variable1=SYSDRIVE
end
item: Set Variable
  Variable=MAINDIR
  Value=temp
  Flags=10000000
end
item: Add Text to INSTALL.LOG
  Text=Variables Set
end
item: Check Disk Space
end
item: Open/Close INSTALL.LOG
end
item: Get Registry Key Value
  Variable=USRDOMAIN
  Key=SYSTEM\CurrentControlSet\Services\Tcpip\Parameters
  Value Name=Domain
  Flags=00000100
end
item: Add Text to INSTALL.LOG
  Text=%USRDOMAIN%
end
item: Get Environment Variable
  Variable=WINDIR
  Environment=windir
end
item: Get Registry Key Value
  Variable=SITE
  Key=SYSTEM\CurrentControlSet\Services\Netlogon\Parameters
  Value Name=DynamicSiteName
  Flags=00000100
end
item: Add Text to INSTALL.LOG
  Text=%SITE%
end
item: Open/Close INSTALL.LOG
  Pathname=%sysdrive%athoc.log
end
item: Add Text to INSTALL.LOG
  Text=Calling .msi installation
end
item: If/While Statement
  Variable=USRDOMAIN
  Value=gdf.aetc.ds.af.mil
  Flags=00000100
end
item: Execute Program
  Pathname=%SYSDRIVE%:\windows\system32\msiexec
  Command Line=/i AtHoc268.msi /qn %ATHOCLOG% %BASEURL% PID=2040619 %SUFFIX% 
  Flags=00001010
end
item: Add Text to INSTALL.LOG
  Text=%USRDOMAIN%
end
item: End Block
end
item: If/While Statement
  Variable=USRDOMAIN
  Value=columbus.aetc.ds.af.mil
  Flags=00000100
end
item: Execute Program
  Pathname=%SYSDRIVE%:\windows\system32\msiexec
  Command Line=/i AtHoc268.msi /qn %ATHOCLOG% %BASEURL% PID=2040592 %SUFFIX% 
  Flags=00001010
end
item: Add Text to INSTALL.LOG
  Text=%USRDOMAIN%
end
item: End Block
end
item: If/While Statement
  Variable=USRDOMAIN
  Value=tyndall.aetc.ds.af.mil
  Flags=00000100
end
item: Execute Program
  Pathname=%SYSDRIVE%:\windows\system32\msiexec
  Command Line=/i AtHoc268.msi /qn %ATHOCLOG% %BASEURL% PID=2040637 %SUFFIX% 
  Flags=00001010
end
item: Add Text to INSTALL.LOG
  Text=%USRDOMAIN%
end
item: End Block
end
item: Get Registry Key Value
  Variable=ATHOC64
  Key=SOFTWARE\wow6432node\AtHocAETC\Desktop
  Value Name=VERSION
  Flags=00000100
end
item: Add Text to INSTALL.LOG
  Text=%athoc%
end
item: If/While Statement
  Variable=ATHOC64
  Value=6.2.11.268
end
item: Add Text to INSTALL.LOG
  Text=HKEY_LOCAL_MACHINE\SOFTWARE\AtHocAETC\desktop\version equals 6.2.11.268 !!!
end
item: Else Statement
end
item: Get Registry Key Value
  Variable=ATHOC32
  Key=SOFTWARE\AtHocAETC\Desktop
  Value Name=VERSION
  Flags=00000100
end
item: If/While Statement
  Variable=ATHOC32
  Value=6.2.11.268
end
item: Add Text to INSTALL.LOG
  Text=HKEY_LOCAL_MACHINE\SOFTWARE\AtHocAETC\desktop\version equals 6.2.11.268 !!!
end
item: Open/Close INSTALL.LOG
  Pathname=%sysdrive%athoc.log
end
item: Exit Installation
  Variable=0
  Flags=1
end
item: End Block
end
item: Add Text to INSTALL.LOG
  Text=Install Completed!
end
item: Open/Close INSTALL.LOG
  Flags=00000001
end
item: End Block
end
item: Exit Installation
  Variable=0
  Flags=0
end
