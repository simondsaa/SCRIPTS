$Computers = Read-Host "PC"
$Array = @()
ForEach ($Computer in $Computers)
{
    If (Test-Connection $Computer -quiet -BufferSize 16 -Ea 0 -count 1 )
    {  
        $SDC = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer).OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation').GetValue('Model')
        $Name = (Get-WmiObject Win32_ComputerSystem -cn $Computer -ErrorAction SilentlyContinue).Name
        $obj = New-Object PSObject
        $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $Name
        $obj | Add-Member -Force -MemberType NoteProperty -Name "SDCVersion" -Value $SDC
        $Array += $obj
    }
}
$Array | Select ComputerName, SDCVersion | OGV -Title "Computer SDCs"