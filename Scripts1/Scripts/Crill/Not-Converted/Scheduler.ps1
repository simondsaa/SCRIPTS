$comp = "52xlwuw3-dgmvv1"
$Create = schtasks.exe /CREATE /TN "Update Stat" /S $comp /SC ONCE /ST 17:00 /RL HIGHEST /RU SYSTEM /TR "powershell.exe -noprofile -File \\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\On_Demand_Stats.ps1" /F
$run = schtasks.exe /RUN /TN "Update Stat" /S $comp
Sleep -Milliseconds 10
$delete  = schtasks.exe /DELETE /TN "Update Stat" /s  $comp /F
