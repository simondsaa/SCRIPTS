$Comp = Read-Host "Name: "
If ($Comp -eq "Foster"){$Compname = "xlwuw-491s64"}
    $Comment = Read-Host "Comment "
    Shutdown /r /f /m \\$Compname /t 20 /c "$Comment"