$computers = Get-Content "C:\Users\1252862141.adm\Desktop\Scripts1\TEST.txt"


foreach ($computer in $computers)
    
    {


$ie = New-Object -ComObject internetexplorer.application
$ie.visible = $true
$ie.navigate("https://www.youtube.com")
#start-sleep -s 5
#$ie.visible = $true

    }

#$chrome = New-Object -ComObject chrome.application
#$chome.visi