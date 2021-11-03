@ECHO OFF
SET MAP="\\tyncesaapspd02\afcesashared$\CEO"
:P
if exist P:\ GOTO Q
net use P: %MAP% /PERSISTENT:YES
GOTO EOF
:Q
if exist Q:\ GOTO X
net use Q: %MAP% /PERSISTENT:YES
GOTO EOF
:X
if exist X:\ GOTO SORRY
net use X: %MAP% /PERSISTENT:YES
GOTO EOF
:SORRY
ECHO Sorry you need to manually map this drive.
ECHO Folder to map: %MAP%
PAUSE
:END