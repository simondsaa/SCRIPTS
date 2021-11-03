$Sec = 300
$Length = $Sec / 100
While($Sec -gt 0)
{
    $Min = [Int](([String]($Sec/60)).split('.')[0])
    $Text = " " + $Min + " minutes " + ($Sec % 60) + " seconds left"
    Write-Progress "Timeout until System Reboot" -Status $Text -PercentComplete ($Sec/$Length)
    Start-Sleep -Seconds 1
    $Sec--
}

Shutdown /r /f /t 10 /c "A system reboot is required to update McAfee Virus Protection on this system, please save all work and your system will reboot in 5 minutes."