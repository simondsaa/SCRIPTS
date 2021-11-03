$names = Get-ChildItem -Path "\\52vkag-fs-netop\NetOps\Vulnerability Management\SoftCertsScan" -Directory
foreach ($name in $names) {
    $name.Name >> "C:\Users\$env:USERNAME\Desktop\PKICertRemoval.txt"
}


$i = 0
$computers = Get-Content "C:\Users\$env:USERNAME\Desktop\PKICertRemoval.txt"
$total = $computers.count
foreach ($computer in $computers) {
    $i++
    Invoke-Command -ComputerName $computer -ScriptBlock {
        $date = Get-Date -Format F
        if (Get-ChildItem -Path c:\ -Include *.p12,*.pfx -File -Recurse -ErrorAction SilentlyContinue) {
            Get-ChildItem -Path c:\ -Include *.p12,*.pfx -File -Recurse -ErrorAction SilentlyContinue | 
            foreach $_ {
                $file = ($_).FullName
                Write-Output "Deleted $file on $date" | Out-File -FilePath "\\52vkag-fs-netop\NetOps\Vulnerability Management\SoftCertsScan\$($computer)\SoftCertsDeleted.txt" -Append
                Remove-Item -Path $file -Force
            }
        }
        else {}
    }
    Write-Progress -Activity "Deleting soft certs --- $i of $total" `
                   -Status "Deleting soft certs on $computer -- $([Math]::Round($i/$total*100))% complete" `
                   -PercentComplete ($i/$total*100)
    Remove-Item -Path "\\52vkag-fs-netop\NetOps\Vulnerability Management\SoftCertsScan\$computer" -Recurse -Force
}