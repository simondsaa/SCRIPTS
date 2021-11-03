<#Code snippets taken from: 

http://stackoverflow.com/questions/26977379/powershell-count-members-of-a-ad-group

#>

Import-Module ActiveDirectory

$pathEmpty = "C:\Temp\groupsEmpty.txt"

Clear-Content $pathEmpty

$Header = `
"Group ID Name" + "|" + `
"Display Name" + "|" + `
"Description" + "|" + `
"Members"


#Write out the header
$Header | Out-File $pathEmpty -Append


$emptys = get-adgroup -properties name, displayname, description, members -Filter name "Java Push Exemption XLWU"  | Select name, displayname, description, members

foreach ($empty in $emptys)
{
#clears previous 
$members = ""
foreach ($member in $empty.members)
  {
    $string = $member.substring(3,$member.indexof(",")-3)
    #$members = $members + ":" + $string
    $string.count 
  }
$listing =`
$empty.Name + "|" + `
$empty.DisplayName + "|" + `
$empty.Description + "|" + `
$members

$listing | Out-File $pathEmpty -Append
}