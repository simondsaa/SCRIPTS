$targetlist = gc "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\HBSS\dat_targets.txt"
$exe = "mcupdate.exe"
$switches = "/update /quiet"
$unavailable = @()
$updatenotran = @()
$failedexe = @()
$updateran = @()

Foreach ($target in $targetlist)
    {
    $ping = New-Object system.net.networkinformation.ping
    $reply = $ping.send($target)

    If ($reply.status -eq "Success")
        {
        
        If((GWmi win32_operatingsystem -computername $target).osarchitecture -eq "32-bit")
            {
            $Path = "C:\Program files\Mcafee\VirusScan Enterprise"        
            $cmd = "$path\$exe $switches"
		    $wmi=([wmiclass]"\\$target\root\cimv2:win32_process")
		    $datprocess = $wmi.create("$cmd")

             If (!$Path)
                {
                "Error in setting '$Path Variable"
                $Updatenotran += $target
                }

             If ($datprocess.returnvalue -ne "0")
                {
                "DAT Update failed to start on $target"
                $failedexe += $target
                }
             If($datprocess.returnvalue -eq "0")
                {
                "DAT Update started on $target"
                $updateran += $target
                }

              }
        
        If((GWmi win32_operatingsystem -computername $target).osarchitecture -eq "64-bit")
            {
            $Path = "C:\Program files (x86)\Mcafee\VirusScan Enterprise"        
            $cmd = "$path\$exe $switches"
		    $wmi=([wmiclass]"\\$target\root\cimv2:win32_process")
		    $datprocess = $wmi.create("$cmd")

             If (!$Path)
                {
                "Error in setting '$Path Variable"
                $Updatenotran += $target
                }

             If ($datprocess.returnvalue -ne "0")
                {
                "DAT Update failed to start on $target"
                $failedexe += $target
                }
             If($datprocess.returnvalue -eq "0")
                {
                "DAT Update started on $target"
                $updateran += $target
                }

              }
        }
    Else
       {
       "$target is unavailable"
       $unavailable += $target
       }
    }

$hash = @{}
$hash.unavailable = $unavailable
$hash.failedexe = $failedexe
$hash.updateran = $updateran
$hash.updatenotran = $updatenotran
$hash | export-clixml "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\HBSS\DAT_Update_Results.xml"