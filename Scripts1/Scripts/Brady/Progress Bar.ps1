$RebootT = 30
$RebootM = $RebootT/60

$Message = "A system reboot is required to update McAfee Virus Protection on this system, please save all work. Your system will reboot in $RebootM minutes."

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

Add-Type -assembly System.Windows.Forms

$Title = "PowerShell Countdown"
$height = 150
$width = 400
$color = "White"

#create the form
$form1 = New-Object System.Windows.Forms.Form
$form1.Text = $title
$form1.Height = $height
$form1.Width = $width
$form1.BackColor = $color

$form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle 
$form1.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

$label1 = New-Object system.Windows.Forms.Label
$label1.Text = "Not started"
$label1.Left = 10
$label1.Top = 70
$label1.Width= $width - 20
$label1.Height = 15
$label1.Font = "Verdana"

$form1.controls.add($label1)

$label2 = New-Object system.Windows.Forms.Label
$label2.Text = $Message
$label2.Left = 10
$label2.Top = 10
$label2.Width= $width - 40
$label2.Height = 70
$label2.Font = "Verdana"

$form1.controls.add($label2)

$progressBar1 = New-Object System.Windows.Forms.ProgressBar
$progressBar1.Name = 'progressBar1'
$progressBar1.Value = 0
#$progressBar1.Style = "Continuous"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = $width - 40
$System_Drawing_Size.Height = 15
$progressBar1.Size = $System_Drawing_Size
$progressBar1.Left = 10
$progressBar1.Top = 90

$form1.Controls.Add($progressBar1)

$form1.Show()| Out-Null

$form1.Focus() | Out-Null

$form1.Refresh()

$Seconds = 1..$RebootT

$i = 0

ForEach ($Sec in $Seconds)
{
    #Write-Host $Sec
    $i++
    [int]$pct = ($i/$Seconds.Count)*100
    $SecLeft = $Seconds.Count - $i
    $Min = [Int](([String]($SecLeft/60)).split('.')[0])
    $progressbar1.Value = $pct
    $label1.text = "$Min" + " minutes " + ($SecLeft % 60) + " seconds left until reboot..."
    $form1.Refresh()

    Start-Sleep -Seconds 1
}

$form1.Close()

#Shutdown /r /f /t 30 /c "TEST"