$Group = Read-Host "Group Name"
write-host " "
write-host $Group
write-host "----------------------------"
ForEach ($User in $Names)
{
    Write-Host $User.Name
}    