$Source = "C:\Users\1392134782A\Desktop\Test1"
$Dest = "C:\Users\1392134782A\Desktop\Test2"
If (!(Test-Path "$Dest"))
{
    New-Item -Path $Dest -Type Directory -Force
}
Get-ChildItem $Source -Recurse | Where {$_.PSIsContainer -eq $false} | ForEach {Copy-Item -Path $_.FullName -Destination $Dest -Force}