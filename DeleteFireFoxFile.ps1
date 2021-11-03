function delete-remotefile {
    PROCESS {
                Write-Host -ForegroundColor Blue -BackgroundColor White "NOTE: * will remove Mozilla Firefox file in Program Files x86."
                $file = "\\$_\c$\Temp\Office 2016"        
                if (test-path $file)
                {
                echo "$_ file exists"
                Remove-Item $file -force -recurse
                echo "$_ file deleted"
                }
            }
}
Get-Content "C:\Temp\MSOffice16_2.txt" | delete-remotefile 0