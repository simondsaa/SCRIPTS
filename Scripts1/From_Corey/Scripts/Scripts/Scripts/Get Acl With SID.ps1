$filename = "Block 1 Test 2"
Get-ChildItem "\\XLWU-FS-003\root\Cons\Shared\Flights\LGC\CC Folder\CDI\Tab C - Background" |
Get-Acl | Select Path -Expand Access | 
Select Path, FileSystemRights, IdentityReference