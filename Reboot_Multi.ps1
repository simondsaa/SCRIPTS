$Start = Get-Date
$Counter = $null

$Path = "C:\Temp\Test.txt"
$Comps = get-content $Path

$ScriptBlock = {
    Param (
    [string]$Comps
    )

ForEach ($Comp in $Comps)
    {
        If (Test-Connection $Comp -quiet -BufferSize 64 -Ea 0 -count 5)
            {
                Shutdown /m \\$Comp /f /r /t 900 /c "Restartin' yer 'puter."
            }

            Else
                {
                    $result = "$Comp is not accessible."
                    $result | Out-File -Verbose C:\Temp\Reboot-Failed.txt -Append
                }  
    }}    

$RunspacePool = [runspacefactory]::CreateRunspacePool(200,200)
$RunspacePool.Open()
$Jobs =
        ForEach ($Comp in $Comps) {

        $Job = [powershell]::Create().
                AddScript($ScriptBlock).
                AddArgument($Comp)
        $Job.RunspacePool = $RunspacePool

        [PSCustomObject]@{
        Pipe = $Job
        Result = $Job.BeginInvoke()
        }
}

$RunspacePool.Close()
$RunspacePool.Dispose()

$Stop = Get-Date
$TimeS = ($Stop - $Start).Seconds
$TimeM = [Math]::Round(($Stop - $Start).TotalMinutes, 0)
Write-Host
Write-Host "Elapsed Time: $TimeM min $TimeS sec" -ForegroundColor Cyan