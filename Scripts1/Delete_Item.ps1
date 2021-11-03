function delete-remotefile {
    PROCESS {
                $file = "\\$_\c$\users\1252862141N\Desktop\TEST.txt"             
                if (test-path $file)
                {
                echo "$_ file exists"
                Remove-Item $file -force -recurse
                echo "$_ file deleted"
                }
            }
}
Get-Content "C:\temp\delete.txt" | delete-remotefile 0