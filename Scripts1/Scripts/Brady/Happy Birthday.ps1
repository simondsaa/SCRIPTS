Filter n {Switch($_){C{262}D{294}E{330}F{349}G{392}A{440}B{415}}}
        #01234567890123456789012345
$Song = "CCDCFECCDCGFDDEDGFECDDEDGF"
$Song[0..3]|n|%{[console]::beep($_,450)}
$Song[4]|n|%{[console]::beep($_,500)}
$Song[5]|n|%{[console]::beep($_,1000)}
Start-Sleep -Milliseconds 20
$Song[6..9]|n|%{[console]::beep($_,450)}
$Song[10]|n|%{[console]::beep($_,550)}
$Song[11]|n|%{[console]::beep($_,1100)}
Start-Sleep -Milliseconds 20
$Song[12..14]|n|%{[console]::beep($_,500)}
$Song[15]|n|%{[console]::beep($_,650)}
$Song[16..19]|n|%{[console]::beep($_,1100)}
#$Song[18]|n|%{[console]::beep($_,1200)}
Start-Sleep -Milliseconds 500
$Song[20..23]|n|%{[console]::beep($_,600)}
$Song[24..25]|n|%{[console]::beep($_,700)}