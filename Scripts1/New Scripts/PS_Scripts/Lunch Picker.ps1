$Exclusions = $null

$Places = "McDonald's", "Wendy's", "Hibachi", "DQ", "Gary's", "Napoli's", "Old Mexico", "Jimmy Johns", "Firehouse", "Chick fil a", "Thai", "BX", "Third Party"

$Lunch = Get-Random -Count 1 $Places
ForEach ($Exclusion in $Exclusions)
{
    If ($Lunch -like $Exclusion)
    {
        Write-Host $Lunch
        $Lunch = Get-Random -Count 1 $Places
    }
}

$c = New-Object -Comobject wscript.shell
$b = $c.popup("$Lunch",0,"Lunch",0)