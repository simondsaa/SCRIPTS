Get-Content C:\Users\timothy.brady\Desktop\Comps.txt |
ForEach {If (Test-Connection $_ -quiet -count 1)
        {write-Host -ForegroundColor Green "$_ is Pingable"}
Else    {write-host -ForegroundColor Red "$_ is Not Pingable"}}