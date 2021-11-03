if exist \\xlwu-fs-003\AFCESAShared\Home\%username% GOTO CHECKUSE
mkdir \\xlwu-fs-003\AFCESAShared\Home\%username%
icacls \\xlwu-fs-003\AFCESAShared\Home\%username% /inheritance:d
icacls \\xlwu-fs-003\AFCESAShared\Home\%username% /remove "tyndall\AFCESA Users"

:CHECKUSE
if exist i:\ GOTO End
net use i: \\xlwu-fs-003\AFCESAShared\Home\%username%

:End