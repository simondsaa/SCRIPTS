$Computers = Get-Content "C:\work\computers.txt"
ForEach ($Computer in $Computers)
{
    If (Test-Connection $Computer -quiet -BufferSize 16 -Ea 0 -count 1)
    {
        write-Host -ForegroundColor Green "$Computer Online"
    }
    Else
    {
        write-host -ForegroundColor Red "$Computer Offline"
    }
}