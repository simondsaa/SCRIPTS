dim objShell, newwallpaper
set objShell = CreateObject("Wscript.Shell")
 
newwallpaper = "C:\Temp\Wallpaper\Yakkety_Yak_Wallpaper.jpg"
 
'If you want to get the current one the use this line
'currentwallpaper = objShell.RegRead ("HKCU\Control Panel\Desktop\Wallpaper")
 
objShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper",newwallpaper, "REG_SZ"