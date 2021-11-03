$Computers = Get-Content "C:\Users\1392134782A\Desktop\BaseComputers.txt"

$GroupSize = 0..59

$Groups = [Math]::Ceiling($Computers.Count/60)

0..$Groups | ForEach {
    
    $Computers[$GroupSize] | ForEach {
    $Name = $_
    Write-Host $Name
    }
    $GroupSize = $GroupSize | ForEach { $_ + 60 }
}