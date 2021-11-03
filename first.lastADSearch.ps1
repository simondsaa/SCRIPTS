$First = Read-Host "First Name"
$Last = Read-Host "Last Name"
$strFilter = "(&(objectCategory=User)(givenName=$First)(sn=$Last))" 
 
$objDomain = New-Object System.DirectoryServices.DirectoryEntry 
 
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher 
$objSearcher.SearchRoot = $objDomain 
$objSearcher.PageSize = 1000 
$objSearcher.Filter = $strFilter 
 
$colProplist = "name" 
foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)} 
 
$colResults = $objSearcher.FindAll() 
 
foreach ($objResult in $colResults) 
    {$objItem = $objResult.Properties; $objItem.name} 
