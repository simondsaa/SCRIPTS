$Message =
"This is a test message.
                          
 Can you see it?"
$c = New-Object -Comobject wscript.shell
$b = $c.popup("$Message",0,"*** Test ***",84)
If ($b -eq 6){Write-Host "You clicked Yes"}
ElseIf ($b -eq 7){Write-Host "You clicked No"}