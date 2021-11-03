[CmdletBinding()]
Param(
 [string[]]$ComputerName = $env:ComputerName,
 [Parameter(mandatory=$true)]
 [string]$Path,
 [string]$OutputFile = "c:\temp\FileVersionInfo.csv"
)
 
$parts = $Path -split ":\\"
$DriveLetter = $parts[0]
$PathwoRoot = $parts[1]
Write-Verbose "Drive letter is $DriveLetter"
Write-Verbose "Remaining Path is $PathwoRoot"
try {
 $Parent = [System.IO.Directory]::GetParent($OutputFile)
 if(Test-Path $Parent.FullName) {
 Write-Verbose "Output folder $($parent.FullName) exists"
 } else {
 throw "Directory $($parent.FullName) not found. Output file cannot be created"
 }
} catch {
 Write-Error "Error occurred while checking output folder. $_"
 return
}
$OutputArr = @()
foreach($Computer in $ComputerName) {
 Write-Host "Querying file version on $Computer"
 $OutputObj = New-Object -TypeName PSobject  
 $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer 
 $OutputObj | Add-Member -MemberType NoteProperty -Name FilePath -Value $Path
 $OutputObj | Add-Member -MemberType NoteProperty -Name FileVersion -Value $null
 $OutputObj | Add-Member -MemberType NoteProperty -Name ProductVersion -Value $null
 $OutputObj | Add-Member -MemberType NoteProperty -Name Status -Value $null
 $OutputObj | Add-Member -MemberType NoteProperty -Name OriginalName -Value $null
 $OutputObj | Add-Member -MemberType NoteProperty -Name FileDescription -Value $null
 $OutputObj | Add-Member -MemberType NoteProperty -Name ProductName -Value $null
 
 if(Test-Connection -ComputerName $Computer -count 1 -quiet) {
 $TargetPath = [string]::format("\\{0}\{1}`$\{2}",$Computer,$DriveLetter,$PathwoRoot)
 Write-verbose "Trying to get $TargetPath file version"
 if(Test-Path $TargetPath) {
 try {
 $VersionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($TargetPath)
 $OutputObj.FileVersion = $VersionInfo.FileVersion
 $OutputObj.ProductVersion = $VersionInfo.ProductVersion
 $OutputObj.OriginalName = $VersionInfo.OriginalName
 $OutputObj.FileDescription = $VersionInfo.FileDescription
 $OutputObj.ProductName = $VersionInfo.ProductName
 $OutputObj.Status = "SUCCESS"
 } catch {
 $OutputObj.Status = "FailedTOQuery"
 }
 } else {
 $OutputObj.Status = "PathNotAccessable"
 }
 } else {
 $OutputObj.Status = "NotReachable"
 }
 $OutputArr += $OutputObj
 Write-Verbose $OutputObj
}
 
$OutputArr | Export-csv $OutputFile -NotypeInformation