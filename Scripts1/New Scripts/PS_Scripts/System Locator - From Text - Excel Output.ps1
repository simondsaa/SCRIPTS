#-----------------------------------------------------------------------------------------#
#                                  Written by SrA Timothy Brady                           #
#                                  Tyndall AFB, Panama City, FL                           #
#-----------------------------------------------------------------------------------------#
$Date = Get-Date -UFormat "%d-%b-%g %H%M"
$domain = "OU=Servers,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL"
$objDomain = [adsi]("LDAP://" + $domain)
$Computers = Get-Content "C:\Users\1392134782A\Desktop\Servers.txt"
$Path = "C:\Users\1392134782A\Desktop\Comp_Loc_$Date.txt"
If (Test-Path $Path){Remove-Item $Path}
ForEach($Computer in $Computers){
    $search = New-Object System.DirectoryServices.DirectorySearcher
    $search.SearchRoot = $objDomain
    $search.Filter = "(&(objectClass=computer)(cn=*$Computer*))"
    $search.SearchScope = "Subtree"
    $search.PageSize = 99999
    $results = $search.FindAll()
    ForEach($item in $results){
        $objComputer = $item.GetDirectoryEntry()
        $Name = $objComputer.cn
        $Loc = $objComputer.Location
        Write-Output "$Name; $Loc" | Out-File $Path -append
        }
    }
$file = “$Path”
$oXL = New-Object -comobject Excel.Application
$oXL.Visible = $true
$oXL.workbooks.OpenText($file,1,1,1,1,$True,$True,$True,$False,$False,$False)

# 1   Tab = True
# 2   Semicolon = True
# 3   Comma = False
# 4   Space = False
# 5   Other = False