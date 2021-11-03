$computers = get-content C:\temp\test.csv
foreach ($c in $computers)
{
if (test-Connection -ComputerName $c -Count 2 -Quiet ) 
        {"$c complete"} 
    else  
        {"$c Failed/No Conn"}
Invoke-Command -ComputerName $c -Scriptblock { [Environment]::SetEnvironmentVariable("JAVA_TOOL_OPTIONS",'-Djava.vendor="Sun Microsystems Inc."',"machine") }  }