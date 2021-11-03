#Find specific file in C:\
Get-ChildItem -Path C:\ -recurse | Where-Object {$_.Name -match 'Bios_Versions.csv'}


#Find named folder
get-childitem -Path PostOOBE -recurse

#Find File type
get-childitem -Path *.csv* -recurse