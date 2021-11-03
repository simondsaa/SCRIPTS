Invoke-Command -ComputerName xlwuw-491s35 -scriptblock {
$obj = new-object -com wscript.shell 
$obj.SendKeys([char]173)}
