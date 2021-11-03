$Comps = Read-Host "Computer Name"
#$Computers = Get-Content C:\Users\1180219788A\Desktop\list.txt
$days = Read-Host "How many days old should the profiles be?"

ForEach ($Comp in $Comps)
    
    { \\"xlwu-fs-05pv\Tyndall_PUBLIC\NCC Admin\Tools\Delprof2 1.6.0\DelProf2.exe" /c:\\$Comp /I /d:$days }
    #{ \\"xlwu-fs-05pv\Tyndall_PUBLIC\NCC Admin\Tools\Delprof2 1.6.0\DelProf2.exe" /c:$Comp /i }