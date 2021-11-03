dir "\\xlwu-fs-03pv\Tyndall_RHS\Shared\CEO" -Directory -recurse | get-acl | 
Where {-NOT $_.AreAccessRulesProtected} | 
Select @{Name="Path";Expression={Convert-Path $_.Path}},AreAccessRulesProtected |
format-table -AutoSize
