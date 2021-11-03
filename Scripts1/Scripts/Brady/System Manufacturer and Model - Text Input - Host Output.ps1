#$Computers = Get-Content "C:\Users\Timothy.Brady\Desktop\Server.txt"
#ForEach($PC in $Computers)
#{
 $PC = "52XLWU-DB-001P"
    If (Test-Connection $PC -quiet -count 1)
    {
    $Model = Get-WmiObject Win32_Computersystem -cn $PC
    Write-Host
    Write-Host "Name         :" $PC
    Write-Host "Manufacturer :" $Model.Manufacturer
    Write-Host "Model        :" $Model.Model
    Write-Host
    }
    Else {Write-Host "$PC not reachable"}
#}