$File = "jre1.7.0_67.msi"

$Deploy = {
   Set-Location "C:\"
   & msiexec.exe /I jre1.7.0_67.msi /quiet /norestart
}

$Computers = Get-Content "C:\Users\timothy.brady\Desktop\Comps.txt"

ForEach ($Computer in $Computers)
{
    Copy-Item -Path \\XLWU-FS-002\Tyndall$\Applications\Java\Java_1.7.0_67\x64\jre1.7.0_67.msi -Destination \\$Computer\C$ -Force
    $Session = New-PSSession -ComputerName $Computer
    Invoke-Command -Session $Session -ScriptBlock $Deploy
    Start-Sleep -Seconds 10
    Remove-PSSession -ComputerName $Computer
    Remove-Item -Path \\$Computer\C$\$File -Force
}