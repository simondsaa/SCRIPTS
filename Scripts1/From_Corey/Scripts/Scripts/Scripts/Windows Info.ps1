Get-Wmiobject Win32_OperatingSystem |
Select Caption, Version, @{Name="Installed On";
Expression={$_.ConvertToDateTime($_.InstallDate)}}