$Computer = "52XLWUW3-DGMVV1"
If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
{
    $RegPath = "SYSTEM\CurrentControlSet\services"        
    $Reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer)
    $RegKey = $Reg.OpenSubKey($RegPath)
    $SubKeys = $RegKey.GetSubKeyNames()
    $Array = @()
    ForEach($SubKey in $SubKeys)
    {
        $Key = $RegPath+"\"+$SubKey 
        $ThisSubKey = $Reg.OpenSubKey($ThisKey)
        $KeyPath = "HKLM:\"+$Key
        If ($Key -like "SYSTEM\CurrentControlSet\services\ac.sharedstore")
        {
             $NewPath = "`"C:\Program Files\Common Files\ActivIdentity\ac.sharedstore.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\ACCMService")
        {
             $NewPath = "`"C:\Program Files (x86)\USAF\ACCM\ACCM_LPC.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\AdobeARMservice")
        {
            $NewPath = "`"C:\Program Files (x86)\Common Files\Adobe\ARM\1.0\armsvc.exe`""
            Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\AdobeFlashPlayerUpdateSvc")
        {
            $NewPath = "C:\Windows\SysWOW64\Macromed\Flash\FlashPlayerUpdateService.exe"
            Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\AeLookupSvc")
        {
             $NewPath = "%systemroot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\ALG")
        {
             $NewPath = "%SystemRoot%\System32\alg.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\AppIDSvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalServiceAndNoImpersonation"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Appinfo")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\AppMgmt")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\aspnet_state")
        {
             $NewPath = "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\aspnet_state.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\AudioEndpointBuilder")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\AudioSrv")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalServiceNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\AxInstSV")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k AxInstSVGroup"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\BDESVC")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\BFE")
        {
             $NewPath = "%systemroot%\system32\svchost.exe -k LocalServiceNoNetwork"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\BITS")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Browser")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\bthserv")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k bthsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\CcmExec")
        {
             $NewPath = "C:\Windows\SysWOW64\CCM\CcmExec.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\CertPropSvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\clr_optimization_v2.0.50727_32")
        {
             $NewPath = "%systemroot%\Microsoft.NET\Framework\v2.0.50727\mscorsvw.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\clr_optimization_v2.0.50727_64")
        {
             $NewPath = "%systemroot%\Microsoft.NET\Framework64\v2.0.50727\mscorsvw.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\clr_optimization_v4.0.30319_32")
        {
             $NewPath = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\mscorsvw.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\clr_optimization_v4.0.30319_64")
        {
             $NewPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\mscorsvw.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\COMSysApp")
        {
             $NewPath = "%SystemRoot%\system32\dllhost.exe /Processid:{02D4B3F1-FD88-11D1-960D-00805FC79235}"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\cphs")
        {
             $NewPath = "%SystemRoot%\SysWow64\IntelCpHeciSvc.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\CryptSvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k NetworkService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\CscService")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\DcomLaunch")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k DcomLaunch"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\defragsvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k defragsvc"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Dhcp")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalServiceNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Dnscache")
        {
            $NewPath = "%SystemRoot%\system32\svchost.exe -k NetworkService"
            Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\dot3svc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\DPS")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalServiceNoNetwork"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\EapHost")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\EFS")
        {
             $NewPath = "%SystemRoot%\System32\lsass.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\EMET_Service")
        {
             $NewPath = "`"C:\Program Files (x86)\EMET 5.1\EMET_Service.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\enstart64")
        {
             $NewPath = "C:\Windows\system32\enstart64.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\enterceptAgent")
        {
             $NewPath = "`"C:\Program Files\McAfee\Host Intrusion Prevention\FireSvc.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\eventlog")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalServiceNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\EventSystem")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Fax")
        {
             $NewPath = "%systemroot%\system32\fxssvc.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\fdPHost")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\FDResPub")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalServiceAndNoImpersonation"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\FontCache")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\FontCache3.0.0.0")
        {
             $NewPath = "%systemroot%\Microsoft.Net\Framework64\v3.0\WPF\PresentationFontCache.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\gpsvc")
        {
             $NewPath = "%windir%\system32\svchost.exe -k GPSvcGroup"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\gupdate")
        {
             $NewPath = "`"C:\Program Files (x86)\Google\Update\GoogleUpdate.exe`" /svc"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\gupdatem")
        {
             $NewPath = "`"C:\Program Files (x86)\Google\Update\GoogleUpdate.exe`" /medsvc"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\hidserv")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\HipMgmt")
        {
             $NewPath = "`"C:\Program Files (x86)\McAfee\Host Intrusion Prevention\HipMgmt.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\hkmsvc")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\HomeGroupListener")
        {
            $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalSystemNetworkRestricted"
            Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\HomeGroupProvider")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalServiceNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\IDriverT")
        {
             $NewPath = "`"C:\Program Files (x86)\Roxio\Roxio MyDVD Basic v9\InstallShield\Driver\1050\Intel 32\IDriverT.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\idsvc")
        {
             $NewPath = "`"%systemroot%\Microsoft.NET\Framework64\v3.0\Windows Communication Foundation\infocard.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\IEEtwCollectorService")
        {
             $NewPath = "%SystemRoot%\system32\IEEtwCollector.exe /V"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\IKEEXT")
        {
             $NewPath = "%systemroot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\IPBusEnum")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\iphlpsvc")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k NetSvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\KeyIso")
        {
             $NewPath = "%SystemRoot%\system32\lsass.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\KtmRm")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k NetworkServiceAndNoImpersonation"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\LanmanServer")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\LanmanWorkstation")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k NetworkService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\lltdsvc")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\lmhosts")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalServiceNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\McAfeeAuditManager")
        {
             $NewPath = "`"C:\Program Files (x86)\McAfee\Policy Auditor Agent\AuditManagerService.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\McAfeeDLPAgentService")
        {
             $NewPath = "`"C:\Program Files\McAfee\DLP\Agent\fcags.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\McAfeeFramework")
        {
             $NewPath = "`"C:\Program Files (x86)\McAfee\Common Framework\FrameworkService.exe`" /ServiceStart"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\McShield")
        {
             $NewPath = "`"C:\Program Files\Common Files\McAfee\SystemCore\\mcshield.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\McTaskManager")
        {
             $NewPath = "`"C:\Program Files (x86)\McAfee\VirusScan Enterprise\vstskmgr.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\mfefire")
        {
             $NewPath = "`"C:\Program Files\Common Files\McAfee\SystemCore\\mfefire.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\mfevtp")
        {
             $NewPath = "`"C:\Windows\system32\mfevtps.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Microsoft SharePoint Workspace Audit Service")
        {
             $NewPath = "`"C:\Program Files (x86)\Microsoft Office\Office14\GROOVE.EXE`" /auditservice"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\MMCSS")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\MozillaMaintenance")
        {
             $NewPath = "`"C:\Program Files (x86)\Mozilla Maintenance Service\maintenanceservice.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\MpsSvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalServiceNoNetwork"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\MSDTC")
        {
             $NewPath = "%SystemRoot%\System32\msdtc.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\MSiSCSI")
        {
             $NewPath = "%systemroot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\msiserver")
        {
             $NewPath = "%systemroot%\system32\msiexec.exe /V"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\napagent")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k NetworkService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Net Driver HPZ12")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k HPZ12"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Netlogon")
        {
             $NewPath = "%systemroot%\system32\lsass.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Netman")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\NetMsmqActivator")
        {
             $NewPath = "`"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\SMSvcHost.exe`" -NetMsmqActivator"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\NetPipeActivator")
        {
             $NewPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\SMSvcHost.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\netprofm")
        {
            $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalService"
            Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\NetTcpActivator")
        {
             $NewPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\SMSvcHost.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\NetTcpPortSharing")
        {
             $NewPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\SMSvcHost.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\NlaSvc")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k NetworkService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\nsi")
        {
             $NewPath = "%systemroot%\system32\svchost.exe -k LocalService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\ose")
        {
             $NewPath = "`"C:\Program Files (x86)\Common Files\Microsoft Shared\Source Engine\OSE.EXE`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\osppsvc")
        {
             $NewPath = "`"C:\Program Files\Common Files\Microsoft Shared\OfficeSoftwareProtectionPlatform\OSPPSVC.EXE`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\p2pimsvc")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalServicePeerNet"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\p2psvc")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalServicePeerNet"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\PcaSvc")
        {
             $NewPath = "%systemroot%\system32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\PeerDistSvc")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k PeerDist"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\PerfHost")
        {
             $NewPath = "%SystemRoot%\SysWow64\perfhost.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\PFERemediation")
        {
             $NewPath = "`"C:\Program Files (x86)\Microsoft PFE Remediation for Configuration Manager\PFERemediation.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\pla")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalServiceNoNetwork"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\PlugPlay")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k DcomLaunch"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Pml Driver HPZ12")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k HPZ12"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\PNRPAutoReg")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalServicePeerNet"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\PNRPsvc")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalServicePeerNet"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\PolicyAgent")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k NetworkServiceNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Power")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k DcomLaunch"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\ProfSvc")
        {
             $NewPath = "%systemroot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\ProtectedStorage")
        {
             $NewPath = "%SystemRoot%\system32\lsass.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\QWAVE")
        {
             $NewPath = "%windir%\system32\svchost.exe -k LocalServiceAndNoImpersonation"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\RasAuto")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\RasMan")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\RemoteAccess")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\RemoteRegistry")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k regsvc"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\RoxMediaDB9")
        {
             $NewPath = "`"C:\Program Files (x86)\Common Files\Roxio Shared\9.0\SharedCOM\RoxMediaDB9.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\rpcapd")
        {
             $NewPath = "`"%ProgramFiles(x86)%\WinPcap\rpcapd.exe`" -d -f `"%ProgramFiles(x86)%\WinPcap\rpcapd.ini`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\RpcEptMapper")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k RPCSS"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\RpcLocator")
        {
             $NewPath = "%SystemRoot%\system32\locator.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\RpcSs")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k rpcss"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SamSs")
        {
             $NewPath = "%SystemRoot%\system32\lsass.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SCardSvr")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalServiceAndNoImpersonation"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SCC_Service")
        {
             $NewPath = "`"C:\Program Files (x86)\SCAP Compliance Checker 3.1.1.1\SCC_Service.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Schedule")
        {
             $NewPath = "%systemroot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SCPolicySvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SDRSVC")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k SDRSVC"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\seclogon")
        {
             $NewPath = "%windir%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SENS")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SensrSvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalServiceAndNoImpersonation"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SessionEnv")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SharedAccess")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\ShellHWDetection")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\smstsmgr")
        {
             $NewPath = "C:\Windows\SysWOW64\CCM\TSManager.exe /service"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SNMP")
        {
             $NewPath = "%SystemRoot%\System32\snmp.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SNMPTRAP")
        {
             $NewPath = "%SystemRoot%\System32\snmptrap.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Spooler")
        {
             $NewPath = "%SystemRoot%\System32\spoolsv.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\sppsvc")
        {
             $NewPath = "%SystemRoot%\system32\sppsvc.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\sppuinotify")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SSDPSRV")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalServiceAndNoImpersonation"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SstpSvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\stisvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k imgsvc"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\stllssvr")
        {
             $NewPath = "`"C:\Program Files (x86)\Common Files\SureThing Shared\stllssvr.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\StorSvc")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\swprv")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k swprv"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\SysMain")
        {
             $NewPath = "%systemroot%\system32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\TabletInputService")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\TapiSrv")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k NetworkService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\TBS")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalServiceAndNoImpersonation"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\TermService")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k NetworkService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Themes")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\THREADORDER")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\TrkWks")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\TrustedInstaller")
        {
             $NewPath = "%SystemRoot%\servicing\TrustedInstaller.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Tumbleweed Desktop Validator")
        {
             $NewPath = "`"C:\Program Files\Tumbleweed\Desktop Validator\DVService.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\UI0Detect")
        {
             $NewPath = "%SystemRoot%\system32\UI0Detect.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\UmRdpService")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\upnphost")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalServiceAndNoImpersonation"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\UxSms")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\VaultSvc")
        {
             $NewPath = "%SystemRoot%\system32\lsass.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\vds")
        {
             $NewPath = "%SystemRoot%\System32\vds.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\VMUSBArbService")
        {
             $NewPath = "`"C:\Program Files (x86)\Common Files\VMware\USB\vmware-usbarbitrator64.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\VSS")
        {
             $NewPath = "%systemroot%\system32\vssvc.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\W32Time")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WatAdminSvc")
        {
             $NewPath = "%SystemRoot%\system32\Wat\WatAdminSvc.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\wbengine")
        {
             $NewPath = "`"%systemroot%\system32\wbengine.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WbioSrvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k WbioSvcGroup"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\wcncsvc")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalServiceAndNoImpersonation"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WcsPlugInService")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k wcssvc"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WdiServiceHost")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WdiSystemHost")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WebClient")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Wecsvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k NetworkService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\wercplsupport")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WerSvc")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k WerSvcGroup"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WinDefend")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k secsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WinHttpAutoProxySvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Winmgmt")
        {
             $NewPath = "%systemroot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WinRM")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k NetworkService"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\Wlansvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\wmiApSrv")
        {
             $NewPath = "%systemroot%\system32\wbem\WmiApSrv.exe"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WMPNetworkSvc")
        {
             $NewPath = "`"%PROGRAMFILES%\Windows Media Player\wmpnetwk.exe`""
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WPCSvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalServiceNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WPDBusEnum")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\wscsvc")
        {
             $NewPath = "%SystemRoot%\System32\svchost.exe -k LocalServiceNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WSearch")
        {
             $NewPath = "%systemroot%\system32\SearchIndexer.exe /Embedding"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\wuauserv")
        {
             $NewPath = "%systemroot%\system32\svchost.exe -k netsvcs"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\wudfsvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalSystemNetworkRestricted"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
        If ($Key -like "SYSTEM\CurrentControlSet\services\WwanSvc")
        {
             $NewPath = "%SystemRoot%\system32\svchost.exe -k LocalServiceNoNetwork"
             Set-ItemProperty -Path $KeyPath -Name "ImagePath" -Value $NewPath -Force
        }
    }
}