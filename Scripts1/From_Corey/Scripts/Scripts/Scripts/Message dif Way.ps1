$CN = Read-Host
$Message = "This is a test"
Invoke-WmiMethod Win32_Process -Name Create -ArgumentList "Msg * $Message" -cn $CN