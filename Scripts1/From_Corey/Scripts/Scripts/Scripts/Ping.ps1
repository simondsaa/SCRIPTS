$Computers = Get-Content "C:\Users\timothy.brady\Desktop\Comps.txt"
ForEach($Computer in $Computers){Ping -n 1 $Computer}