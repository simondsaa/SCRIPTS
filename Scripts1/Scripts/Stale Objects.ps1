# The 60 is the number of days from today since the last logon.

$then = (Get-Date).AddDays(-90)
Get-ADComputer -Property Name,lastLogonDate -Filter {lastLogonDate -lt $then} -SearchBase 'OU=Tyndall AFB Computers,OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL' |
 Select-Object Name | Export-Csv C:\Users\1180219788A\Desktop\stale.csv
