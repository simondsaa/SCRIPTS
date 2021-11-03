@ECHO OFF
if %PROCESSOR_ARCHITECTURE%==x86 (
if exist "C:\Program Files\Mozilla Firefox\firefox.exe" GOTO End
\\Xlwu-fs-002\tyndall$\Applications\Mozilla Firefox 17.0\FireFox-Setup-17.0.1.exe /S /v/qn
ECHO Firefox Installed...
) else (
if exist "C:\Program Files (x86)\Mozilla Firefox\firefox.exe" GOTO End
\\Xlwu-fs-002\tyndall$\Applications\Mozilla Firefox 17.0\FireFox-Setup-17.0.1.exe /S /v/qn
ECHO Firefox Installed...
)
:End