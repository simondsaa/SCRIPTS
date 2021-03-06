$spath = "C:\Users\timothy.brady\Desktop"
$opath = "C:\Users\timothy.brady\Documents\Results.csv"
#-----------------------------------------------------------
$SSN_Regex = "[0-9]{3}[-| ][0-9]{2}[-| ][0-9]{4}"
$PN_Regex = "[0-9]{3}[-| ][0-9]{3}[-| ][0-9]{4}"
#-----------------------------------------------------------
Get-ChildItem -Path $spath -Recurse -Filter *.docx| Select-String -Pattern $SSN_Regex | Select-Object Path,Filename,Matches | Export-CSV $opath
Get-ChildItem -Path $spath -Recurse -Filter *.docx| Select-String -Pattern $PN_Regex | Select-Object Path,Filename,Matches | Export-CSV $opath