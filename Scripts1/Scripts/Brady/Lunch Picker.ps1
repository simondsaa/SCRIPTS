$Exclusions = "JR's", "Napoli's", "Chick fil a"

$Places = "McDonald's", "Wendy's", "Hibachi", "DQ", "Gary's", "Napoli's", "Old Mexico", "Jimmy Johns", "Firehouse", "Chick fil a", "Thai", "BX", "Third Party", "JR's", "Bowling"

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

$Body = "Lunchinator 5000 says:

$Lunch"

Send-MailMessage -From "timothy.brady.11@us.af.mil" -To "lonnie.stringer.3@us.af.mil", "leej.tobler@us.af.mil", "ashley.thompson.21@us.af.mil" -Priority High -Subject "Lunch" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil