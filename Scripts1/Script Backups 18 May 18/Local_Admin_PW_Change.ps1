$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
exit $LASTEXITCODE
}

$computers = Read-Host "Enter the Computer Name" 
# Update username / password as needed
$username = "usaf_admin"
$password = "zaq1XSW@zaq1XSW@"

# Lists to store success / failed attempts
$success = New-Object System.Collections.Generic.List[string]
$failure = New-Object System.Collections.Generic.List[string]

# Loop through each computer
foreach ($computer in $computers) {
    # Attempt to change the password on the computer, ignoring any errors
    try {
        ([ADSI] "WinNT://$computer/$username").SetPassword("$password")
    } catch {}
    # On success:
    if ($?) { 
        $success.Add($computer)
        Write-Host "Success: $computer" -ForegroundColor Green
    }
    # On failure:
    else { 
        $failure.Add($computer)
        Write-Host "Failure: $computer" -ForegroundColor Red
    }
}

# Uncomment to export results to a file:
#$success | Out-File "C:\Users\1383807847.adm\desktop\success.txt"
#$failure | Out-File "C:\Users\1383807847.adm\desktop\failure.txt"

