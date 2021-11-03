$SCCM = New-Object -ComObject UIResource.UIResourceMgr
$SCCM.ExecuteProgram("`*","INE0068C",$true)