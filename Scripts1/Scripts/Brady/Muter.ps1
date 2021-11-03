$c = New-Object -Comobject wscript.shell
$b = $c.popup("Good day sir!",0,"Hello",80)

$obj = New-Object -Comobject wscript.shell
$a = $obj.SendKeys([char]173)