$SCCM = New-Object -ComObject UIResource.UIResourceMgr
$SCCM.GetAvailableApplications() | Select ID, PackageID, PackageName | Format-List
#$SCCM.ExecuteProgram("ID","PackageId",$true)