$Value = "C:\Program Files (x86)\Java\jre7"
$current = [Environment]::GetEnvironmentVariable("Java_home","Machine")

IF ($current -ne $Value) {[Environment]::SetEnvironmentVariable("Java_home", "$Value", "Machine")} ELSE {exit}
