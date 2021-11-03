Invoke-Command -ComputerName server01 -ScriptBlock { 
    Start-Process c:\windows\temp\installer.exe -ArgumentList '/silent' -Wait
}