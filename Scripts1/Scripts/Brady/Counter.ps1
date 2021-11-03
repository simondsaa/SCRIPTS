$Orgs = Get-Content -Path "C:\work\Computers.txt"
ForEach ($Org in $Orgs)
{
    $Match = Select-String -Path "C:\work\computers.txt" -Pattern $Org -AllMatches | Select -ExpandProperty Matches
    $Match
}