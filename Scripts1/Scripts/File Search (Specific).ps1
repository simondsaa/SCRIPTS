#---------------------------------------------------------------------------------
#                          Written by Robie_0n3
#                          Tyndall AFB, Panama City, FL
#---------------------------------------------------------------------------------

$SSN_Regex = "[0-9]{3}[-| ][0-9]{2}[-| ][0-9]{4}"
$PN_Regex = "[0-9]{3}[-| ][0-9]{3}[-| ][0-9]{4}"

$filePath = "\\xlwu-fs-01pv\Tyndall_ANG"
Get-ChildItem -Include *.pst, *.tmp, *.wav, *.jpg, *.wma, *.mp3, *.mpeg, *.exe, *.avi -Recurse -Force $filePath -ErrorAction SilentlyContinue | where {!$_.PSIsContainer} | 
Select-Object Name, Directory, Length, CreationTime, LastAccessTime, LastWriteTime | Export-Csv C:\Users\1180219788A\Desktop\Results\FileSearch_FS01pv_ANG.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $SSN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS01pv_ANG.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $PN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS01pv_ANG.csv

$filePath = "\\xlwu-fs-02pv\Tyndall_44_FG"
Get-ChildItem -Include *.pst, *.tmp, *.wav, *.jpg, *.wma, *.mp3, *.mpeg, *.exe, *.avi -Recurse -Force $filePath -ErrorAction SilentlyContinue | where {!$_.PSIsContainer} | 
Select-Object Name, Directory, Length, CreationTime, LastAccessTime, LastWriteTime | Export-Csv C:\Users\1180219788A\Desktop\Results\FileSearch_FS02pv_44FG.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $SSN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS02pv_44FG.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $PN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS02pv_44FG.csv


$filePath = "\\xlwu-fs-03pv\Tyndall_NCOA"
Get-ChildItem -Include *.pst, *.tmp, *.wav, *.jpg, *.wma, *.mp3, *.mpeg, *.exe, *.avi -Recurse -Force $filePath -ErrorAction SilentlyContinue | where {!$_.PSIsContainer} | 
Select-Object Name, Directory, Length, CreationTime, LastAccessTime, LastWriteTime | Export-Csv C:\Users\1180219788A\Desktop\Results\FileSearch_FS03pv_NCOA.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $SSN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_NCOA.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $PN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_NCOA.csv

$filePath = "\\xlwu-fs-03pv\Tyndall_RHS"
Get-ChildItem -Include *.pst, *.tmp, *.wav, *.jpg, *.wma, *.mp3, *.mpeg, *.exe, *.avi -Recurse -Force $filePath -ErrorAction SilentlyContinue | where {!$_.PSIsContainer} | 
Select-Object Name, Directory, Length, CreationTime, LastAccessTime, LastWriteTime | Export-Csv C:\Users\1180219788A\Desktop\Results\FileSearch_FS03pv_RHS.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $SSN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_RHS.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $PN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_RHS.csv

$filePath = "\\xlwu-fs-04pv\Tyndall_53_WEG"
Get-ChildItem -Include *.pst, *.tmp, *.wav, *.jpg, *.wma, *.mp3, *.mpeg, *.exe, *.avi -Recurse -Force $filePath -ErrorAction SilentlyContinue | where {!$_.PSIsContainer} | 
Select-Object Name, Directory, Length, CreationTime, LastAccessTime, LastWriteTime | Export-Csv C:\Users\1180219788A\Desktop\Results\FileSearch_FS03pv_53WEG.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $SSN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_53WEG.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $PN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_53WEG.csv

$filePath = "\\xlwu-fs-04pv\Tyndall_325_FW"
Get-ChildItem -Include *.pst, *.tmp, *.wav, *.jpg, *.wma, *.mp3, *.mpeg, *.exe, *.avi -Recurse -Force $filePath -ErrorAction SilentlyContinue | where {!$_.PSIsContainer} | 
Select-Object Name, Directory, Length, CreationTime, LastAccessTime, LastWriteTime | Export-Csv C:\Users\1180219788A\Desktop\Results\FileSearch_FS03pv_325FW.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $SSN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_325FW.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $PN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_325FW.csv

$filePath = "\\xlwu-fs-04pv\Tyndall_325_MSG"
Get-ChildItem -Include *.pst, *.tmp, *.wav, *.jpg, *.wma, *.mp3, *.mpeg, *.exe, *.avi -Recurse -Force $filePath -ErrorAction SilentlyContinue | where {!$_.PSIsContainer} | 
Select-Object Name, Directory, Length, CreationTime, LastAccessTime, LastWriteTime | Export-Csv C:\Users\1180219788A\Desktop\Results\FileSearch_FS03pv_325MSG.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $SSN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_325MSG.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $PN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_325MSG.csv

$filePath = "\\xlwu-fs-04pv\Tyndall_325_MXG"
Get-ChildItem -Include *.pst, *.tmp, *.wav, *.jpg, *.wma, *.mp3, *.mpeg, *.exe, *.avi -Recurse -Force $filePath -ErrorAction SilentlyContinue | where {!$_.PSIsContainer} | 
Select-Object Name, Directory, Length, CreationTime, LastAccessTime, LastWriteTime | Export-Csv C:\Users\1180219788A\Desktop\Results\FileSearch_FS03pv_325MXG.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $SSN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_325MXG.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $PN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_325MXG.csv

$filePath = "\\xlwu-fs-04pv\Tyndall_325_OG"
Get-ChildItem -Include *.pst, *.tmp, *.wav, *.jpg, *.wma, *.mp3, *.mpeg, *.exe, *.avi -Recurse -Force $filePath -ErrorAction SilentlyContinue | where {!$_.PSIsContainer} | 
Select-Object Name, Directory, Length, CreationTime, LastAccessTime, LastWriteTime | Export-Csv C:\Users\1180219788A\Desktop\Results\FileSearch_FS03pv_325OG.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $SSN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_325OG.csv
Get-ChildItem -Path $filePath -Exclude *.dll, *.exe -Recurse | Select-String -Pattern $PN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Results\PII_Search_FS03pv_325OG.csv