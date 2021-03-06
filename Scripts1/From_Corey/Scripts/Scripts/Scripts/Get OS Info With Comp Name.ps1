Get-Wmiobject -cn TYNTS01 Win32_OperatingSystem |
    Select @{Name="Operating System"; Expression={$_.Caption}}, Version,
    @{Name="Installed On"; Expression={$_.ConvertToDateTime($_.InstallDate)}}, Organization,
    @{Name="Computer Name"; Expression={$_.CSName}}