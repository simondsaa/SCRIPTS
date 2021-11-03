$servers= get-content "C:\Users\1394844760A\Desktop\Scripting Test Bed\names.txt"
$output = "C:\Users\1394844760A\Desktop\Scripting Test Bed\LocalAdmin.csv"

foreach($server in $servers)
{
    $group =[ADSI]"WinNT://$server/Administrators" 
    $members = @($group.Invoke("Members")) | %{$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)}
}
ForEach ($user in $members)
{
     Write-Host $user
}