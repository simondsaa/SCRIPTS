$GroupNames = Get-Content "C:\Users\timothy.brady\Desktop\Groups.txt"

ForEach ($Name in $GroupNames)
{
    $Groups = Get-ADGroup -filter {Name -like $Name} | Select-Object Name

    ForEach ($Group in $Groups)
    {   write-host " "
        write-host "$($group.name)"
        write-host "----------------------------"

        $Names = Get-ADGroupMember -identity $($group.name)
        ForEach ($User in $Names.Name)
        {
            Write-Host $User
        }
    }
}