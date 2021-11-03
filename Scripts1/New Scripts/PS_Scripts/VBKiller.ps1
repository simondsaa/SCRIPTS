$Computers = Get-Content C:\Users\timothy.brady\Desktop\Computers2.txt
ForEach ($Computer in $Computers)
{
    $Files = "Msjet35.dll"
    ForEach ($File in $Files)
    {
        If ($File -ne $null)
        {
            REGSVR32 /U /S "\\$Computer\C$\Windows\System32\$File"
            Remove-Item -Path "\\$Computer\C$\Windows\System32\$File" -Force
        }
    }
}