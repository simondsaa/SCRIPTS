$code = @'
using System;
using System.Runtime.InteropServices;

namespace CloseButtonToggle {
  internal static class WinAPI {
    [DllImport("kernel32.dll")]
    internal static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool DeleteMenu(IntPtr hMenu,
                           uint uPosition, uint uFlags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool DrawMenuBar(IntPtr hWnd);

    [DllImport("user32.dll")]
    internal static extern IntPtr GetSystemMenu(IntPtr hWnd,
               [MarshalAs(UnmanagedType.Bool)]bool bRevert);

    const uint SC_CLOSE     = 0xf060;
    const uint MF_BYCOMMAND = 0;

    internal static void ChangeCurrentState(bool state) {
      IntPtr hMenu = GetSystemMenu(GetConsoleWindow(), state);
      DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND);
      DrawMenuBar(GetConsoleWindow());
    }
  }

  public static class Status {
    public static void Disable() {
      WinAPI.ChangeCurrentState(false); //its 'true' if need to enable
    }
  }
}
'@

Add-Type $code
[CloseButtonToggle.Status]::Disable()

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