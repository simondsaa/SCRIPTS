############################################
#Script to Update McAfee on a local machine#
#Author A1C Klinger, Samuel (SME)          #
#Created 14 Feb 2018                       #
############################################
 If ((Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\McAfee\Agent").AgentVersion -like "5.0.6*"){}

 Else {
       If (Test-Path -Path "C:\Program Files\McAfee\Agent\x86\FrmInst.exe"){
       Invoke-Command -ScriptBlock{
                                   cd "C:\Program Files\McAfee\Agent\x86\"
                                   .\FrmInst.exe /forceuninstall /silent}     
       Start-Sleep -Seconds 600
       New-Item -Path C:\ -Name McAfee -ItemType container -Force
       Copy-Item -Path "\\vkag-fs-01pv\SeymourJohnson_4FW_SJ_ALL\sj_all_csa_info\Software Patches\HBSS & McAfee AV\FramePkg.exe" -Destination "C:\McAfee\"
       Invoke-Command -ScriptBlock {
                                    cd "C:\McAfee\"
                                    .\FramePkg.exe /install=agent /forceinstall /silent}
       Start-Sleep -Seconds 600
       Remove-Item -Path "C:\McAfee" -Recurse -Force
       $env:COMPUTERNAME + ": system updated" >>  "\\52vkag-fs-netop\NetOps\Vulnerability Management\McAfee\Upgraded_Systems.txt"
       Invoke-Command -ScriptBlock {
                                    GPUpdate /force
                                    }
       }
        
        Else {New-Item -Path C:\ -Name McAfee -ItemType container -Force
              Copy-Item -Path "\\vkag-fs-01pv\SeymourJohnson_4FW_SJ_ALL\sj_all_csa_info\Software Patches\HBSS & McAfee AV\FramePkg.exe" -Destination "C:\McAfee\"
              Invoke-Command -ScriptBlock {
                                   cd "C:\McAfee\"
                                   .\FramePkg.exe /install=agent /forceinstall /silent}
              Start-Sleep -Seconds 600
              Remove-Item -Path "C:\McAfee" -Recurse -Force
              $env:COMPUTERNAME + ": system updated" >>  "\\52vkag-fs-netop\NetOps\Vulnerability Management\McAfee\Upgraded_Systems.txt"
              Invoke-Command -ScriptBlock {
                                    GPUpdate /force
                                          }      
              }
}