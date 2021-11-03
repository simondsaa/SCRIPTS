$Servers = Get-Content C:\Users\timothy.brady\Desktop\Server.txt
ForEach ($Name in $Servers)
{
    New-Item -ItemType Directory -Path "\\XLWU-FS-002\Tyndall$\Server Remediation\SCAP\Results\$Name"
}