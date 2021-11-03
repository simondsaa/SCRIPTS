@ECHO OFF
@TITLE Map Drives
net use M: "\\xlwu-fs-001\AFNORTH_Media" \PERSISTENT:YES
cls
net use P: "\\xlwu-fs-001\pfps$" \PERSISTENT:YES
cls
net use S: "\\xlwu-fs-001\ANG$\Shared" \PERSISTENT:YES
cls
net use T: "\\xlwu-fs-002\Tyndall$" \PERSISTENT:YES
cls
@TITLE Drives Mapped Successfully...Please close this window.
ECHO Drives Mapped Successfully...Please close this window.
pause
