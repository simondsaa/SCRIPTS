#This is the path you want the results to be saved
$Path = "C:\Users\Timothy.Brady\Desktop\Programs.xml"

#This is the path containing your list of Computers
$Computers = Get-Content "C:\Users\Timothy.Brady\Desktop\Comps.txt"
$Array = @()
ForEach($PC in $Computers){
    If (Test-Connection $PC -quiet -count 1){
        $RegPath = "Software\\McAfee"
        $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$PC)
        $ThisKey = $RegPath+"\\AVEngine"
        $ThisSubKey = $Reg.OpenSubKey($ThisKey)
            $obj = New-Object PSObject
            $obj | Add-Member -Force -MemberType NoteProperty -Name "ComputerName" -Value $PC
            $obj | Add-Member -Force -MemberType NoteProperty -Name "DatDate" -Value $($ThisSubKey.GetValue("AVDatDate"))
            $obj | Add-Member -Force -MemberType NoteProperty -Name "DatVersion" -Value $($ThisSubKey.GetValue("AVDatVersion"))
            $Array += $obj}
    Else {Write-Host "$PC Not reachable"}
}
$Array | Select ComputerName, DatDate, DatVersion | Export-CliXML $Path
$Result = Read-Host "Would you like to see your results now? Y or N"
If($Result -eq "Y"){Import-CliXML $Path | OGV} ElseIf($Result -eq "N"){Write-Host "Data has been saved here: $Path"}