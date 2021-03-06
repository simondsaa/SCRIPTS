$spath = "\\52XLWUW3-DKPVV1\C$\Users\timothy.brady\Desktop"
$opath = "C:\Users\timothy.brady\Desktop\Results.txt"
#-----------------------------------------------------------
$SSN_Regex = "[0-9]{3}[-| ][0-9]{2}[-| ][0-9]{4}"
$PN_Regex = "[0-9]{3}[-| ][0-9]{3}[-| ][0-9]{4}"
#-----------------------------------------------------------
Get-ChildItem $spath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $SSN_Regex | Select-Object Path,Filename,Matches | Format-List -Force | Out-File $opath
Get-ChildItem $spath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $PN_Regex | Select-Object Path,Filename,Matches | Format-List -Force | Out-File $opath -Append