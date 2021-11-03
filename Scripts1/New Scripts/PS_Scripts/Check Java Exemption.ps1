$Computers = Get-Content \\xlwu-fs-05pv\Tyndall_PUBLIC\Patching\Java\JavaTargets.txt
ForEach ($Computer in $Computers)
{
    $Groups = (Get-ADPrincipalGroupMembership (Get-ADComputer $Computer).DistinguishedName).Name
    If ($Groups -like "Java Push Exemption*")
    {
        Write-Host "$Computer is a member of Java Exemption Group"
        "$Computer" | Out-File \\xlwu-fs-05pv\Tyndall_PUBLIC\Patching\Java\JavaContacts.txt -Append -Force
    }
    If (!($Groups -like "Java Push Exemption*"))
    {
        Write-Host "$Computer can be patched"
        "$Computer" | Out-File \\xlwu-fs-05pv\Tyndall_PUBLIC\Patching\Java\JavaPatch.txt -Append -Force
    }
}