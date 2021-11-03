#$Shares = Get-Content C:\Users\1180219788A\Desktop\FileShares1.txt
$Shares = "\\xlwu-fs-05pv\Tyndall_PUBLIC"

Foreach ($Share in $Shares)
    {Get-ChildItem -Path $Share -Recurse |
     Select-Object DirectoryName, BaseName, Length | Export-Csv -Path C:\Users\1180219788A\Desktop\Share_Listing.csv -Encoding ASCII -NoTypeInformation}