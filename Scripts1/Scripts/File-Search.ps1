$Start = Get-Date
$Counter = $null

$InputFile = Get-Content "C:\Users\1180219788A\Desktop\shares.txt"
$SSN_Regex = "[0-9]{3}[-| ][0-9]{2}[-| ][0-9]{4}"
$PN_Regex = "[0-9]{3}[-| ][0-9]{3}[-| ][0-9]{4}"

$ScriptBlock = {
    Param (
    [string]$InputFile
    )

    ForEach ($FilePath in $InputFile) {

         If ($FilePath -match "Tyndall_ANG"){($Output = "FileSearch_FS01pv_ANG")}
     ElseIf ($FilePath -match "Tyndall_44_FG"){($Output = "FileSearch_FS02pv_44FG")}
     ElseIf ($FilePath -match "Tyndall_NCOA"){($Output = "FileSearch_FS03pv_NCOA")}
     ElseIf ($FilePath -match "Tyndall_RHS"){($Output = "FileSearch_FS04pv_RHS")}
     ElseIf ($FilePath -match "Tyndall_53_WEG"){($Output = "FileSearch_FS04pv_53WEG")}
     ElseIf ($FilePath -match "Tyndall_325_FW"){($Output = "FileSearch_FS04pv_325FW")}
     ElseIf ($FilePath -match "Tyndall_325_MSG"){($Output = "FileSearch_FS04pv_325MSG")}
     ElseIf ($FilePath -match "Tyndall_325_MXG"){($Output = "FileSearch_FS04pv_325MXG")}
     ElseIf ($FilePath -match "Tyndall_325_OG"){($Output = "FileSearch_FS04pv_325OG")}
           
Get-ChildItem -Include *.pst, *.tmp, *.wav, *.jpg, *.wma, *.mp3, *.mpeg, *.exe, *.avi -Recurse -Force $FilePath -ErrorAction SilentlyContinue | where {!$_.PSIsContainer} | 
    Select-Object Name, Directory, Length, CreationTime, LastAccessTime, LastWriteTime | Export-Csv C:\Users\1180219788A\Desktop\Searches\$Output.csv
}}

$ScriptBlockB = {
    Param (
    [string]$InputFile,[string]$SSN_Regex,[string]$PN_Regex
    )

        ForEach ($FilePath in $InputFile) {

         If ($FilePath -match "Tyndall_ANG"){($Output2 = "PIISearch_FS01pv_ANG")}
     ElseIf ($FilePath -match "Tyndall_44_FG"){($Output2 = "PIISearch_FS02pv_44FG")}
     ElseIf ($FilePath -match "Tyndall_NCOA"){($Output2 = "PIISearch_FS03pv_NCOA")}
     ElseIf ($FilePath -match "Tyndall_RHS"){($Output2 = "PIISearch_FS04pv_RHS")}
     ElseIf ($FilePath -match "Tyndall_53_WEG"){($Output2 = "PIISearch_FS04pv_53WEG")}
     ElseIf ($FilePath -match "Tyndall_325_FW"){($Output2 = "PIISearch_FS04pv_325FW")}
     ElseIf ($FilePath -match "Tyndall_325_MSG"){($Output2 = "PIISearch_FS04pv_325MSG")}
     ElseIf ($FilePath -match "Tyndall_325_MXG"){($Output2 = "PIISearch_FS04pv_325MXG")}
     ElseIf ($FilePath -match "Tyndall_325_OG"){($Output2 = "PIISearch_FS04pv_325OG")}

Get-ChildItem -Path $FilePath -Exclude *.dll, *.exe -Recurse -Force | Select-String -Pattern $SSN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Searches\$Output2.csv
#Get-ChildItem -Path $FilePath -Exclude *.dll, *.exe -Recurse -Force | Select-String -Pattern $PN_Regex | Select-Object Filename,Path | Export-CSV C:\Users\1180219788A\Desktop\Searches\$Output2.csv
}}

$RunspacePool = [runspacefactory]::CreateRunspacePool(1,1)
$RunspacePool.Open()
$Jobs =
        ForEach ($FilePath in $InputFile) {

        $Job =  [powershell]::Create().
                AddScript($ScriptBlock).
                AddArgument($FilePath)

        $Job.RunspacePool = $RunspacePool

        [PSCustomObject]@{
        Pipe = $Job
        Result = $Job.BeginInvoke()
        }}
   
        ForEach ($FilePath in $InputFile) {

        $Job =  [powershell]::Create().
                AddScript($ScriptBlockB).
                AddArgument($FilePath).
                AddArgument($SSN_Regex).
                AddArgument($PN_Regex)
        
        $Job.RunspacePool = $RunspacePool

        [PSCustomObject]@{
        Pipe = $Job
        Result = $Job.BeginInvoke()
        }}

Do {
   $Counter++
   cls
   Write-Host 'Please Wait.  Scanning...'
   Start-Sleep -Seconds 1
} While ( $Jobs.Result.IsCompleted -contains $false)

$RunspacePool.Close()
$RunspacePool.Dispose()

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan