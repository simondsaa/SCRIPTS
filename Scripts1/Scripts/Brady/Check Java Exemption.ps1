$Computers = Get-Content C:\Users\1392134782A\Desktop\Comps.txt
ForEach ($Computer in $Computers)
{
    $Groups = (Get-ADPrincipalGroupMembership (Get-ADComputer $Computer).DistinguishedName).Name
    If ($Groups -like "Java Push Exemption*")
    {
        Write-Host "$Computer is a member of Java Exemption Group"
    }
    If (!($Groups -like "Java Push Exemption*"))
    {
        Write-Host "$Computer can be patched"
        "$Computer" | Out-File "C:\Users\1392134782A\Desktop\JavaPatch.txt" -Append -Force
    }
}