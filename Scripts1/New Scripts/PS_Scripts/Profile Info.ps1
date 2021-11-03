$Computer = "xlwuw-491s35"
$ProfPath = "\\$Computer\C$\Users"
$Profiles = Get-ChildItem $ProfPath

$AdminProf = 0

ForEach ($Profile in $Profiles)
{
    If ($Profile -like "*.adm")
    {
        $AdminProf += 1
    }
    
    $Files = Get-ChildItem $ProfPath\$Profile -Recurse -ErrorAction SilentlyContinue
    $FileSize = $Files | Measure-Object -Property length -Sum
    $SizeMB = "{0:N1}" -f ($FileSize.sum/1MB)
    
    Write-Host "$Profile   - $SizeMB MB"
}

Write-Host "Number of Profiles :" $Profiles.Count"total | $AdminProf admin profile(s)" 