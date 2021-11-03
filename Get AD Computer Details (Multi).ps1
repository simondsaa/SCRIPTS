$Path = Read-Host "PC Names"
$Computers = Get-Content $Path


ForEach ($Computer in $Computers){


$ADComputerProperties = @(`
"Operatingsystem",
"OperatingSystemServicePack",
"Created",
"Enabled",
"LastLogonDate",
"IPv4Address",
"Location"
)
 
$SelectADComputerProperties = @(`
"Name",
"OperatingSystem",
"OperatingSystemServicePack",
"Created",
"Enabled",
"LastLogonDate",
"IPv4Address",
"Location"
)

Get-ADComputer -Filter * -Properties $ADComputerProperties | 
select $SelectADComputerProperties | Export-Csv -Path C:\Temp\Logs\1.csv -Encoding ASCII
}

