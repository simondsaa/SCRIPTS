$Computers = Get-Content "C:\Users\timothy.brady\Desktop\Computers.txt"
ForEach ($Computer in $Computers)
{
    If (Test-Connection $Computer -quiet -BufferSize 16 -Ea 0 -count 1)
    {    
        $Info = Get-WmiObject -ComputerName $Computer Win32_ComputerSystem -ErrorAction SilentlyContinue
        Write-Host
        Write-Host $Info.Name":"$Info.SystemType
    }
    Else
    {
        Write-Host
        Write-Host "$Computer : unavailable"
    }
}