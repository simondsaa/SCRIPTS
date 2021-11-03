$OPath = "C:\Users\timothy.brady\Desktop\Results.txt"
$RootFolder = Get-ChildItem -Recurse -Path "\\XLWU-FS-004\root\325 FW\325 MSG\325 CS"
ForEach ($Folder in $RootFolder)
{
$Access = Get-ACL $Folder.FullName 
    If($Access.Owner -ne "BUILTIN\Administrators")
    {
    Write-OutPut "File Name: " $Folder.FullName "Owner: " $Access.Owner |
    Out-File $OPath -append
    } 
 }