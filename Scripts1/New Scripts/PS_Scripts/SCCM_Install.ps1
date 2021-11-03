$ID = "*"
$PackageID = "INE005D1"

$Install = New-Object -ComObject UIResource.UIResourceMgr
$Install.ExecuteProgram($ID, $PackageID, $true)