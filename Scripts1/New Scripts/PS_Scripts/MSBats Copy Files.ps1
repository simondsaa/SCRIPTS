$Computers = Get-Content "C:\Users\timothy.brady\Desktop\Comps.txt"
$Files = Get-ChildItem "C:\MSBats"
ForEach ($Computer in $Computers)
{
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        ForEach ($File in $Files)
        {
            CD C:\MSBats
            Copy-Item $File \\$Computer\C$ -Force
        }
    }
}