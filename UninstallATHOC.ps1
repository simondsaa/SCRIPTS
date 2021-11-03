$app = Get-WmiObject -Class Win32_Product | Where-Object { 
    $_.Name -match "AtHocAMC"}

$app.Uninstall()

$app = Get-WmiObject -Class Win32_Product | Where-Object { 
    $_.Name -match "AtHocUSAF"}

$app.Uninstall()

$app = Get-WmiObject -Class Win32_Product | Where-Object { 
    $_.Name -match "AtHocAFMC"}

$app.Uninstall()

$app = Get-WmiObject -Class Win32_Product | Where-Object { 
    $_.Name -match "AtHocGOV"}

$app.Uninstall()