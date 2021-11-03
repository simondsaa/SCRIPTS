$Task = "False"
$OSInfo = Get-WmiObject Win32_OperatingSystem 
Try {
          IF ($OSInfo.OSArchitecture -eq "64-bit") {
              Start-process "C:\Program Files (x86)\McAfee\VirusScan Enterprise\mcupdate.exe" /update
               $Task = "True"
          } ELSE {
              start-process "C:\Program Files\McAfee\VirusScan Enterprise\mcupdate.exe" /update
               $Task = "True"
              }

}

Catch { $Task = "Errored"}
"$env:ComputerName : $Task : $($OSInfo.OSArchitecture) : $((get-date).DateTime)" | out-file -append "C:\Users\1180219788A\Desktop\DAT_Update.txt"
