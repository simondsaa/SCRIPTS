Write-Host -ForegroundColor Blue -BackgroundColor White "NOTE: Typing * will remove all files in Temp folder."
$item = Read-Host "File + Extension you want removed from remote Temp folder"

function delete-remotefile {
    PROCESS {
                $file = "\\$Comp\c$\Temp\$item"             
                if (test-path $file)
                {
                echo "$Comp file exists"
                Remove-Item $file -force -recurse
                echo "$Comp file deleted"
                }
            }
}
 
 $Comp = Read-Host "PC Name" | delete-remotefile 0
