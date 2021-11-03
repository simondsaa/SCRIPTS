$UserIdentityList=@(
    Invoke-Expression 'query.exe user /server:localhost' | # Query user sessions
    Select-Object -Skip 1 |                                # Skip header row
    ForEach {$_.Trim() -replace '\s\s.+',''} |             # Extract usernames
    ForEach {
        try {                                              # Attempt to translate
            $username = $_
            [Security.Principal.WindowsIdentity]$_ 
            }
            catch {
            "$username could not be identified"
            }
            }
            )