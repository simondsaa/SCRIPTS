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
 
Get-ADComputer -Searchbase "OU=Tyndall AFB Computers,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" -Filter * -Properties $ADComputerProperties  |  `
select $SelectADComputerProperties | OGV -Title ComputerDetails 