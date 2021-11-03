$Computers = Get-Content "C:\Users\timothy.brady\Desktop\Comps.txt"
$Path = "HKLM:\SOFTWARE\Policies\Microsoft\ccmsetup"
$Key = "SetupParameters"
ForEach($Computer in $Computers)
{
    $SiteCode = Get-ItemProperty -path $Path -name $Key | Select SetupParameters
    If ($SiteCode  -like "*XLW*")
    {
        Write-Host "$Computer"
    }
}    