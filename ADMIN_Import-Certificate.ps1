# Filepath to source certificate to import (.cer, .sst, etc.)
    $CertPath = "C:\temp\PIV.cer"

#Filepath to computer list
    $computers = Get-Content "C:\Temp\test.txt"

# Trusted Root Store Path:
    $TrustedRoot = "Cert:\LocalMachine\Root"

# Intermediate Certificate Path:
    $Intermediate = "Cert:\LocalMachine\CA"

# Local Machine Path:
    $LocalMachine = "Cert:\LocalMachine\My"

#############################################################################

foreach ($computer in $computers)
{
    if (Test-Connection $computer -Count 1 -ea 0 -Quiet)
    { 
        Write-Host "CHECKING $computer..." -ForegroundColor Green
        Import-Certificate -FilePath "$CertPath" -CertStoreLocation $TrustedRoot -Verbose
    } 
    else 
    { 
        Write-Host "$computer Is Unreachable" -ForegroundColor Red
 
    }
  
}