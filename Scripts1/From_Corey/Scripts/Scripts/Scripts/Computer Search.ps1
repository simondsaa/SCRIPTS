$Systems = Get-Content "C:\Users\timothy.brady\Desktop\Comps.txt"
ForEach ($System in $Systems)
{$Name = Get-WmiObject Win32_ComputerSystem -cn $System -ErrorAction SilentlyContinue
    If ($Name.UserName -like "*AREA52\1392134782A*"){ 
        Write-Host "Name: " $Name.Name
        Write-Host "User: " $Name.UserName}
    }