$Computers = Get-Content C:\Users\timothy.brady\Desktop\Comps.txt
ForEach ($Computer in $Computers)
{
    REG ADD "\\$Computer\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v AllowRemoteRPC /t REG_DWORD /d 1 /f
}