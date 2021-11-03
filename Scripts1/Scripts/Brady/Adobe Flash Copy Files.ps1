$Computers = Get-Content "C:\Users\timothy.brady\Desktop\Comps.txt"
$Files = Get-ChildItem "C:\Viper"
$Path = "\\$Computer\C$\Users\Public"
ForEach ($Computer in $Computers)
{
    ForEach ($File in $Files)
    {
        CD C:\Viper
        Copy-Item $File $Path -Force
    }
}