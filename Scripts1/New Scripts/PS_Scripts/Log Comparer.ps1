$LogPath1 = "C:\Users\1392134782A\Desktop\Log1.txt"

$LogPath2 = "C:\Users\1392134782A\Desktop\Log2.txt"

$Log1 = (Get-Content -Path $LogPath1).Split(" ")

$Log2 = (Get-Content -Path $LogPath2).Split(" ")

Compare-Object -ReferenceObject $Log1 -DifferenceObject $Log2 | Select InputObject