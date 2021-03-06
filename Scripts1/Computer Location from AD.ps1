$domain = "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$objDomain = [adsi]("LDAP://" + $domain)
$Computers = Get-Content "C:\Users\1252862141.adm\Desktop\Scripts1\Enable_Local_Admin.txt"
#$Computers = Read-Host "Computer Name"
ForEach($computer in $Computers)
{
    $search = New-Object System.DirectoryServices.DirectorySearcher
    $search.SearchRoot = $objDomain
    $search.Filter = "(&(objectClass=computer)(cn=*$Computer*))"
    $search.SearchScope = "Subtree"
    $results = $search.FindAll()
    ForEach($item in $results){
        $objComputer = $item.GetDirectoryEntry()
        $Name = $objComputer.cn
        $Loc = $objComputer.Location
        Write-Output "$Name; $Loc" | Out-File C:\work\Tyndall_Loc.txt -append
            }
        }