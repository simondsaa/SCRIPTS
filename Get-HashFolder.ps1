<#
Script evaluates Staged content on afnet boxes to ensure they have not been tampered with.. 
A new key is created for the purpose of reporting at:

HKLM:\Software\USAF\SDC with a name of "PostOOBEValidation" and there are 3 possible values:
    
    0 - Content is valid
    1 - Content has been tampered with
    2 - Content has not been staged on the client
#>


param(
  [String] $Folder = "C:\Servicing\PostOOBE",
  [String] $File   = "C:\ProgramData\setupcomplete.cmd",
  [String] $File2  = "C:\Users\Default\AppData\Local\Microsoft\Windows\WSUS\SetupConfig.ini"
)

function Set-SystemOnlyPermissions($Item)
{
    $Acl = Get-Acl $Item
    $rules = $acl.access | Where { $_.IdentityReference -ine "NT Authority\System" -and $_.IdentityReference -ine "BUILTIN\Administrators" }
    ForEach($rule in $rules) {
        cacls $($Item) /e /r $($rule.IdentityReference) | Out-Null
    }
 
}

function Fix-SetupConfig($Item)
{
#Set knowns
$WSUSExist = Test-Path C:\Users\Default\AppData\Local\Microsoft\Windows\WSUS
$SetupConfigFile = Test-Path $Item
$header = "[SetupConfig]"
$noreboot = "NoReboot"
$PostOOBE = 'PostOOBE="C:\ProgramData\SetupComplete.cmd"'
$InstallDrivers = 'InstallDrivers="C:\Servicing\Drivers"'

#Build out SetupConfig.ini
    If($WSUSExist -eq $false)
    
        {
        Write-Host -ForegroundColor Red "WSUS Directory not present. Creating..."
        New-Item -ItemType Directory -Path `
        C:\Users\Default\AppData\Local\Microsoft\Windows\WSUS
        }
    
    Else{
        Write-Host -ForegroundColor Green "WSUS Directory exist. Continuning..."
        }                                
        
    If($SetupConfigFile -eq $true)
    
        {
        Write-Host -ForegroundColor Yellow "SetupConfig.ini already exist. Overwriting with known good version."
        Remove-Item -Path C:\Users\Default\AppData\Local\Microsoft\Windows\WSUS\SetupConfig.ini -Force
        }

    Write-Host -ForegroundColor Red "Writing to SetupComplete.ini..."
    New-Item -ItemType File -Path `
    C:\Users\Default\AppData\Local\Microsoft\Windows\WSUS\SetupConfig.ini
    Add-Content -Value $header -Path C:\Users\Default\AppData\Local\Microsoft\Windows\WSUS\SetupConfig.ini
    Add-Content -Value $noreboot -Path C:\Users\Default\AppData\Local\Microsoft\Windows\WSUS\SetupConfig.ini
    Add-Content -Value $InstallDrivers -Path C:\Users\Default\AppData\Local\Microsoft\Windows\WSUS\SetupConfig.ini
    Add-Content -Value $PostOOBE -Path C:\Users\Default\AppData\Local\Microsoft\Windows\WSUS\SetupConfig.ini
}
Function Get-FolderHash($folder){
   
   $files = Get-ChildItem "$folder" -Recurse -File -include *.csv, *.ps1, *.exe, *.cmd, *.mof, *.ini
   
   $allBytes = @()
   foreach ($file in $files)
   {
       $allBytes += Get-Content $file.FullName -Encoding Byte
   }
   $hasher = [System.Security.Cryptography.SHA256]::Create()
   $ret = [string]::Join("",$($hasher.ComputeHash($allBytes) | %{"{0:x2}" -f $_}))
   return $ret
}

if((Test-Path $Folder) -eq $false)
{
    Write-host "Invalid Folder: $folder" -ForegroundColor Red
    Write-Output "Content is not staged on this client."
    Set-ItemProperty "HKLM:\Software\USAF\SDC" -Name PostOOBEValidation -Value 2
    Exit 0
}

$Hash = Get-FolderHash $folder


$CompareHash = "2c6eee55d23f89068273294c40ccb93cbe08b03a2a927a90b379a81be5ed2bb0"

$CompareHash = $CompareHash.Replace("`r`n","")

#PostOOBE Compare
if($Hash -eq $CompareHash)
{
    Write-Output "Staged content is valid!"
    Write-Output "Setting Reg.."
    Set-ItemProperty "HKLM:\Software\USAF\SDC" -Name PostOOBEValidation -Value 0
}
else
{
    Write-Output "Folder content is not valid!"
    Write-Output "Hash returned:" $Hash
    Write-Output "Setting PostOOBEValidation to 1"
    Set-ItemProperty "HKLM:\Software\USAF\SDC" -Name PostOOBEValidation -Value 1
}

Set-SystemOnlyPermissions $Folder


if((Test-Path $File) -eq $false)
{
    Write-Output "Can't Find File: $File"
    exit -2
}
$FileHash = Get-FolderHash $File

$CompareFileHash = "5D1931F29DF1B8AC0A5C3E19F8D2A1B1D7F3F6283BFCEE28EC8567D44A5791E8"

if($FileHash -eq $CompareFileHash)
{
    Write-Output "SetupComplete is valid!"
}
else
{
    Write-Output "SetupComplete is not valid! Fixing.."
    Write-Output $FileHash

    $Validate = Get-ItemProperty "HKLM:\Software\USAF\SDC" -Name PostOOBEValidation -ErrorAction SilentlyContinue
    if($Validate -eq 0)
    {
        if((Test-Path "$Folder\setupcomplete.cmd") -eq $true)
        {
            Copy-Item "$Folder\setupcomplete.cmd" $File -Force
        }
    }

}

Set-SystemOnlyPermissions $File


if((Test-Path $File2) -eq $false)
{
    Write-Output "Can't Find File: $File2"
    exit -3
}
$File2Hash = Get-FolderHash $File2

$CompareFile2Hash = "21E7FBBF0F0140827C9BBA0A30E9D6AD92D7592B07620834F527C46D3FC3863F"

if($File2Hash -eq $CompareFile2Hash)
{
    Write-Output "SetupConfig is valid!"
}
else
{
    Write-Output "SetupConfig is not Valid! Fixing..."
    Write-Output $File2Hash

    #Replace contents
    Fix-SetupConfig $File2
}

#Set Permissions
Set-SystemOnlyPermissions($File2)