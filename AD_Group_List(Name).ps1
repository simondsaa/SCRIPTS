$Identity = Read-Host "Enter Group as listed in AD"
get-adgroupmember -Identity "$Identity" -Recursive | select name