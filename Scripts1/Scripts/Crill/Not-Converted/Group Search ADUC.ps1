CLS
#Continue until closed
While ($true) {
#Clear out values
$objName = $null
$objGroupSearch = $null
Write-Host "Enter your AD Group name that you would like to search for" -Foregroundcolor Green
#Set group name
$objName = Read-Host "Group Name"
#Search for group
$objGroupSearch = Get-ADGroupMember -Identity $objName
#Display in a readable fashion
$objGroupSearch.name | sort-object
Write-Host "Total Group Members: $(($objGroupSearch.name).count)" -ForegroundColor Yellow
}
