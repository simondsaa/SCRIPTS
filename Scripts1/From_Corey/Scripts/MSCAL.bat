@ECHO OFF
@TITLE MSCAL.OCX FILE FIX

REM MSCAL.OCX FILE REGISTER SCRIPT by C. Jarrett, 101 ACOMS CS, Tyndall AFB, DSN 742-0272

REM *** Copies OCX File to System32 ***
xcopy "\\xlwu-fs-001\ang$\Shared\_03 AOC\ACOMS\SCO\SCOC\CSA\SOFTWARE\MSCAL.OCX" "C:\Windows\System32\" /Y


RegSvr32 MSCAL.OCX /s