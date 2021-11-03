cscript %windir%\system32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r 131.55.28.116 -h 131.55.28.116 -o raw
if %PROCESSOR_ARCHITECTURE%==x86 (
%windir%\system32\rundll32.exe printui.dll,PrintUIEntry /if /b "CPD Lexmark X740 Series PS3" /f "S:\_03 AOC\ACOMS\SCO\SCOC\CSA\DRIVERS\Printer and Scanner Drivers\Lexmark_X740_Series_ADO_Win_32_PS\Drivers\sysPS32Win\Drivers\Print\GDI\LMADON40.inf" /r "131.55.28.116" /m "Lexmark X740 Series PS3"
 ) else (
%windir%\system32\rundll32.exe printui.dll,PrintUIEntry /if /b "CPD Lexmark X740 Series PS3" /f "S:\_03 AOC\ACOMS\SCO\SCOC\CSA\DRIVERS\Printer and Scanner Drivers\Lexmark_X740_Series_ADO_Win_32_PS\Drivers\sysPS32Win\Drivers\Print\GDI\LMADON40.inf" /r "131.55.28.116" /m "Lexmark X740 Series PS3" ) 