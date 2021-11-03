# Open a Telnet window. You need to install Telnet services first.
# example: psexec \\xlwul-42093d dism /online /enable-feature /featurename:TelnetClient

# Run the keystrokes
Invoke-Command -ComputerName xlwul-42093d -Scriptblock {
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.SendKeys]::('{CAPSLOCK}')}
