$Computers = Get-Content "C:\Temp\2.txt"
ForEach ($Computer in $Computers)
{
    If (Test-Connection $Computer -quiet -BufferSize 16 -Ea 0 -count 1)
    {
        write-Host -ForegroundColor Green "$Computer Online"
    }
    Else
    {
        write-output $Computer | out-file C:\Temp\SDC_NOT_Pinging.txt
        write-host -ForegroundColor DarkGreen "$Computer Offline"
    }
}