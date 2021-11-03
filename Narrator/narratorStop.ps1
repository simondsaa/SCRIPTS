$p = Get-Process "narrator"
stop-process -InputObject $p -force