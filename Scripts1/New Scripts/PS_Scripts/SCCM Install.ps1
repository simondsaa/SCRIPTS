$SCCM = New-Object -ComObject UIResource.UIResourceMgr
$SCCM.GetAvailableApplications() | Select ID, PackageID, PackageName | Where {$_.PackageName -like "*"} | Format-List
#$SCCM.ExecuteProgram("`*","INE0050A",$true)