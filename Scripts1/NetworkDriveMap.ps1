net use S: \\xlwu-fs-01pv\Tyndall_ANG\Shared

net use T: \\xlwu-fs-05pv\Tyndall_PUBLIC

net use P: \\xlwu-fs-02pv\Tyndall_PFPS

Add-Type -AssemblyName PresentationCore,PresentationFramework
$ButtonType = [System.Windows.MessageBoxButton]::Ok
$MessageIcon = [System.Windows.MessageBoxImage]::Asterisk
$MessageBody = "Your share drives have been mapped. Brought to you by TSgt Simonds. This will not map your personal drive. Please feel free to 'share' (pun intended)."
$MessageTitle = "Network Drive Mapping Status"
 
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
 
Write-Host "Your choice is $Result"

