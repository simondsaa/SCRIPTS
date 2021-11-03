Robocopy.exe "\\tyncesaapsql02\d$\Backup" "G:\Backups$\SQL Server Backup" *.* /SEC /MIR /r:1 /w:1 /NP /LOG:C:\Robocopy\SQL_DB.txt
