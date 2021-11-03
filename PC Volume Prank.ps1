﻿#Hide Volume icon in taskbar:  Xlwuw-759074\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer - DWORD HideSCAVolume -value 1 (to hide)
#To do this, enable remote registry:  SET:  Set-service -name RemoteRegistry -StartupType Automatic / START:  Get-Service -name RemoteRegistry | start-service
#stop/start explorer:  STOP:  stop-process -processname explorer -force / START:  start-process explorer
#Enter PSSession
#run this script
#adjust volume with [audio]::Volume = 0.2 (Equals 20%, 1 Equals 100%)
#[audio]::Mute = $false or $true
#Set-service -name RemoteRegistry -StartupType Automatic
#Get-Service -name RemoteRegistry | start-service
#$RegKey = “HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer”
#if (-Not(Test-Path “$RegKey”)) {
#    New-Item -Path $RegKey -Force
#}
#Set-ItemProperty -Path $RegKey -Name “HideSCAVolume” -Type Dword -Value 1
#stop-process -processname explorer -force
#sleep -Seconds 5
#start-process explorer
Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume
{
    // f(), g(), ... are unused COM method slots. Define these if you care
    int f(); int g(); int h(); int i();
    int SetMasterVolumeLevelScalar(float fLevel, System.Guid pguidEventContext);
    int j();
    int GetMasterVolumeLevelScalar(out float pfLevel);
    int k(); int l(); int m(); int n();
    int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, System.Guid pguidEventContext);
    int GetMute(out bool pbMute);
}
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice
{
    int Activate(ref System.Guid id, int clsCtx, int activationParams, out IAudioEndpointVolume aev);
}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator
{
    int f(); // Unused
    int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }
public class Audio
{
    static IAudioEndpointVolume Vol()
    {
        var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
        IMMDevice dev = null;
        Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(/*eRender*/ 0, /*eMultimedia*/ 1, out dev));
        IAudioEndpointVolume epv = null;
        var epvid = typeof(IAudioEndpointVolume).GUID;
        Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
        return epv;
    }
    public static float Volume
    {
        get { float v = -1; Marshal.ThrowExceptionForHR(Vol().GetMasterVolumeLevelScalar(out v)); return v; }
        set { Marshal.ThrowExceptionForHR(Vol().SetMasterVolumeLevelScalar(value, System.Guid.Empty)); }
    }
    public static bool Mute
    {
        get { bool mute; Marshal.ThrowExceptionForHR(Vol().GetMute(out mute)); return mute; }
        set { Marshal.ThrowExceptionForHR(Vol().SetMute(value, System.Guid.Empty)); }
    }
}
'@

[audio]::Mute = $false
sleep -Seconds 30
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1
sleep -Milliseconds 250
[audio]::Volume = 0
sleep -Milliseconds 500
[audio]::Volume = 0.7
sleep -Milliseconds 750
[audio]::Volume = 0.3
sleep -Milliseconds 500
[audio]::Volume = 0
sleep -Milliseconds 250
[audio]::Volume = 1