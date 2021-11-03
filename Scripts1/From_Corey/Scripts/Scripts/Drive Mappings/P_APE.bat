if exist P:\ GOTO Q
:P
net use P: "\\tyncesaapspd02\afcesashared$\CC" /PERSISTENT:YES
:Q
net use Q: "\\tyncesaapspd02\afcesashared$\CC" /PERSISTENT:YES
:END