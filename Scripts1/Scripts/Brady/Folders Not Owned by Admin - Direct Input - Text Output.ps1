$OPath = "C:\Users\1180219788A\Desktop\Bad_Owner_Results.txt"
$RootFolder = Get-ChildItem -Recurse -Path "\\xlwu-fs-03pv\Tyndall_RHS"
ForEach ($Folder in $RootFolder)
{
$Access = Get-ACL $Folder.FullName 
    If($Access.Owner -ne "BUILTIN\Administrators")
    {
    Write-OutPut "File Name: " $Folder.FullName "Owner: " $Access.Owner |
    Out-File $OPath -append
    } 
 }