#SCCM Validation
#Returns True or False if computername exists in certificate and saves in array
$Computers = "52xlwuw3-dkpvv2"

$valid = $null
$invalid = $null
$Valid = @()
$InValid = @()




#SCCM Repair
    Foreach ($Computer in $Computers) {
    $check = certutil.exe -verifystore \\$computer\SMS 1
    $check
    $filter = $check -match $Computer
        IF ($check -match "Issuer: CN=SMS, CN=$computer" ) {
        Write-Host "$($Computer) is valid" -ForegroundColor Green
        $Valid += $Computer
            }
    ELSE { 
           $del1 = certutil.exe -delstore \\$Computer\SMS 1
           $del0 = certutil.exe -delstore \\$Computer\SMS 0
            Write-Host "$($Computer) is not Valid" -ForegroundColor yellow
            $Invalid += $Computer
            $del1
            $del0
            Stop-Service CCMEXEC -PassThru
            Write-Host "Start Sleep"
            Start-Sleep -Seconds 3
            Start-Service CCMEXEC -PassThru
        } 
  }
  Write-Host "Valid: $($Valid.count)" -ForegroundColor DarkYellow
  Write-Host "Invalid: $($InValid.count)" -ForegroundColor DarkYellow
