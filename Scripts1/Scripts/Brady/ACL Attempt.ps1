$Script:FoldersFound = New-Object System.Collections.ArrayList
$Directory = "\\xlwu-fs-004\325 CS$\SCO\SCOO"
$GroupName = "DLS_325 CS_SCO"
$Folders = Get-ChildItem $Directory -Recurse | Where-Object {$_.PSIsContainer}
ForEach ($Folder in $Folders)
{
    $ACL = Get-ACL -Path $Folder.FullName
    ForEach ($ACE in $ACL.access)
    {
        If (($ACE.IdentityReference) -like "*\$GroupName")
        {
            $FolderFound.Add($Folder.Fullname) | Out-Null
            Break
        }
    }
}
$FoldersFound.Sort()
ForEach($Folder in $FoldersFound)
{
    Write-Host $Folder
}