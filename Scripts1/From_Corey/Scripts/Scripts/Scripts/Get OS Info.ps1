Get-Wmiobject -cn TYNDALLNP-APP Win32_OperatingSystem |
    Select @{Name="Operating System"; Expression={$_.Caption}},
    Version,
    @{Name="Installed On"; Expression={$_.ConvertToDateTime($_.InstallDate)}},
    Organization,
    @{Name="Computer"; expression={$_.CSName}} | Out-File C:\Users\timothy.brady\Desktop\Test.txt