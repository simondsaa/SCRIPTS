$Server = "TYNCSSCBWKT0603"
$Process = [WMICLASS]"\\$TYNCSSCBWKT0603\ROOT\CIMV2:win32_process"
$Result = $Process.Create("C:\Brady's Stuff\Batch Files\Message With IF.bat")