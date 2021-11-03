$Comp = read-host "PC Name"
icm -computername $Comp -scriptblock {msg * Message}
