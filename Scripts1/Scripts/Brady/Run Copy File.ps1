$NetPath = "\\XLWU-FS-004\325 CS`$\325 CS Shared\CCRI - Lt Mayers\Remediation\Patches"
$file1 = "32BITPatches.bat"
$file2 = "64BITPatches.bat"

$Computers = Get-Content "C:\Users\timothy.brady\Desktop\Comps.txt"
ForEach ($Computer in $Computers)
{
    Function Copy-Files
    {
        Copy-Item "$NetPath\$file1" -Destination \\$Computer\C$
        Copy-Item "$NetPath\$file2" -Destination \\$Computer\C$
    }
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        Invoke-Command -ComputerName $Computer -ScriptBlock {Copy-Files}
    }
}