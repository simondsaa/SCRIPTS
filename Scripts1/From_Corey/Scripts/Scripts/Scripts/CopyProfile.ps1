        $comp = Read-Host "PC with profile"
        $EDI = Read-Host "EDI + Designator"
        $comp2 = Read-Host "PC needing profile"

        $Users = Get-ChildItem \\$Comp\C$\Users | Where-Object { $EDI -contains $_.Name }
        ForEach ($User in $Users)
        {
            $Desktop = "\\$Comp\C$\Users\$EDI\Desktop"
            $Documents = "\\$Comp\C$\Users\$EDI\Documents"
            $Favorites = "\\$Comp\C$\Users\$EDI\Favorites"
            $Downloads = "\\$Comp\C$\Users\$EDI\Downloads"
            $Music = "\\$Comp\C$\Users\$EDI\Music"
            $Pictures = "\\$Comp\C$\Users\$EDI\Pictures"
            $Videos = "\\$Comp\C$\Users\$EDI\Videos"
            $Destination = "\\$Comp2\C$\Temp\ProfileBackups\$EDI"
            Copy-Item $Desktop "$Destination\Desktop" -Recurse -Force
            Copy-Item $Documents "$Destination\Documents" -Recurse -Force
            Copy-Item $Favorites "$Destination\Favorites" -Recurse -Force
            Copy-Item $Downloads "$Destination\Downloads" -Recurse -Force
            Copy-Item $Music "$Destination\Music" -Recurse -Force
            Copy-Item $Pictures "$Destination\Pictures" -Recurse -Force
            Copy-Item $Videos "$Destination\Videos" -Recurse -Force
        }
    
{
$End = Get-Date
$TimeS = ($End - $Start).Seconds
$TimeM = ($End - $Start).Minutes
Write-Host "Run Time: $TimeM min $TimeS sec"
}