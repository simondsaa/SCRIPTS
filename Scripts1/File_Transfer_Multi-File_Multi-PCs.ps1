$a = Get-Content "C:\work\File_Transfer_Multi_Computers.txt"  
 
foreach ($i in $a)  
 
{$files= get-content "C:\work\File_Transfer_Multi_Files.txt" 
foreach ($file in $files) 
{Copy-Item $file -Recurse -Destination \\$i\C$\Users\public\Desktop -force} 
} 