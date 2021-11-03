$Start = Get-Date


$Computer = "xlwul-42093d"
#$Computers = Get-Content "C:\work\computers.txt"
#ForEach ($Computer in $Computers)
#{
    If (!(Test-Path \\xlwu-fs-004\root\Profiles\$Computer))
    {
        New-Item -ItemType Directory -Path "\\xlwuw-491s35\C$\temp\temp" -Force
    }
    Sleep -Seconds 1
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        $Users = Get-ChildItem \\$Computer\C$\Users
        ForEach ($User in $Users)
        {
            $Desktop = "\\$Computer\C$\Users\$User\Desktop"
            $Documents = "\\$Computer\C$\Users\$User\Documents"
            $Favorites = "\\$Computer\C$\Users\$User\Favorites"
            $Downloads = "\\$Computer\C$\Users\$User\Downloads"
            $Music = "\\$Computer\C$\Users\$User\Music"
            $Pictures = "\\$Computer\C$\Users\$User\Pictures"
            $Videos = "\\$Computer\C$\Users\$User\Videos"
            $Destination = "\\xlwu-fs-004\root\Profiles\$Computer\Users\$User"
            Copy-Item $Desktop "$Destination\Desktop" -Recurse -Force
            Copy-Item $Documents "$Destination\Documents" -Recurse -Force
            Copy-Item $Favorites "$Destination\Favorites" -Recurse -Force
            Copy-Item $Downloads "$Destination\Downloads" -Recurse -Force
            Copy-Item $Music "$Destination\Music" -Recurse -Force
            Copy-Item $Pictures "$Destination\Pictures" -Recurse -Force
            Copy-Item $Videos "$Destination\Videos" -Recurse -Force
        }
    }
#}

$End = Get-Date
$TimeS = ($End - $Start).Seconds
$TimeM = ($End - $Start).Minutes
Write-Host "Run Time: $TimeM min $TimeS sec"