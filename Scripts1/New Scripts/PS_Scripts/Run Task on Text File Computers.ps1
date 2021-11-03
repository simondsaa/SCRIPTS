$Files = Get-ChildItem "\\xlwu-fs-05pv\Tyndall_PUBLIC\Stats\Profile Cleanup"
ForEach ($File in $Files)
{
    If ($File.PSIsContainer -eq $false)
    {
        $Comps = $File.Name.Split(".")[0]
        ForEach ($Comp in $Comps)
        {
            $Run = schtasks.exe /RUN /TN "Delete Old Profiles" /S $Comp
        }
    }
}