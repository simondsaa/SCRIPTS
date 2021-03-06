$Computers = Get-Content "C:\Users\1252862141.adm\Desktop\Scripts1\Enable_Local_Admin.txt"

ForEach ($Computer in $Computers)
{
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        $User = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer
        If ($User.UserName -ne $null)
        {
            $EDI = $User.UserName.TrimStart("AREA52\")
            $UserName = (Get-ADUser "$EDI" -Properties DisplayName).DisplayName
        }
        Else
        {
            $UserName = "No user logged on"
        }
    }
    
    Write-Host "$Computer   - $UserName"
}