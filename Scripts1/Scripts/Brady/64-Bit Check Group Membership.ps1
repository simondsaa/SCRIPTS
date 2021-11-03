$Computers = Get-Content \\xlwu-fs-004\Home\1392134782A\Desktop\Comps.txt
ForEach ($Computer in $Computers)
{
    $Groups = (Get-ADPrincipalGroupMembership (Get-ADComputer $Computer).DistinguishedName).Name
    If ($Groups -like "GLS_TYNDALL_SDC3.xOSDx64")
    {
        Write-Host "$Computer is a member of GLS_TYNDALL_SDC3.xOSDx64"
    }
    If (!($Groups -like "GLS_TYNDALL_SDC3.xOSDx64"))
    {
        Write-Host "$Computer can be patched"
        "$Computer" | Out-File "\\xlwu-fs-004\Home\1392134782A\Desktop\Add.txt" -Append -Force
    }
}