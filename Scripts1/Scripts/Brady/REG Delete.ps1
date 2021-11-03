$Computers = Get-Content "C:\Users\timothy.brady\Desktop\Comps.txt"
ForEach($Computer in $Computers)
{
    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(‘LocalMachine’,$Computer)
    $SK = $Reg.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",$True)
    $SK.deletesubkey("{4A03706F-666A-4037-7777-5F2748764D10}")
}