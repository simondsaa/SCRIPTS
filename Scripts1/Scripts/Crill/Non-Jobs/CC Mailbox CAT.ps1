###
#
# Mailbox CAT scanner for CC's
# SSgt Crill, Christian 325 CS/SCOO
#
#
###


cls
$ErrorActionPreference = "SilentlyContinue"

$users =get-aduser -Filter {ObjectClass -eq "User"} -searchbase "OU=Tyndall AFB, OU=AFCONUSEAST,OU=Bases,DC=area52,dc=afnoapps,dc=usaf,dc=mil"  -Properties DisplayName, CN, o, Office, ExtensionAttribute5| ? {$_.office -like "CC"} |select -Property *
   
   $array = @()
Foreach ($user in $users) {   
$obj = [PSCustomObject] @{
            Name = $user | select -expandproperty DisplayName 
            Organization = $user.o[0]
            Office= $user.Office
            MailboxCAT = $user.ExtensionAttribute5
            }
$array += $obj
            }        

"CC Designated Users that are not CAT I"

$array | where {$_.mailboxCAT -ne "1"} |Format-Table -AutoSize
$array | where {$_.mailboxCAT -ne "1"} |Format-Table -AutoSize | Export-csv C:\CC_MailboxCAT.csv

         
         
         
         
         
         

