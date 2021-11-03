$OPath = "C:\Users\1252862141.adm\Desktop\Scripts1\Pop.txt"
$RootFolder = Get-ChildItem -Recurse -Path "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CS\SCO\SCOO"
ForEach ($Folder in $RootFolder){
$Access = Get-ACL $Folder.FullName 
    If($Access.Owner -ne "BUILTIN\Administrators"){
    Write-OutPut "File Name: " $Folder.FullName "Owner: " $Access.Owner | 
    Out-File $OPath -append 
    }
}