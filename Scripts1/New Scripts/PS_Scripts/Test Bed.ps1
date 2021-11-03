$Start = Get-Date

#$Group = "GLS_TYNDALL_SDC3.xOSDx64"

#$Computers = (Get-ADGroupMember -Identity $Group).Name
$Computers = Get-Content "\\xlwu-fs-004\Home\1392134782A\Desktop\Comps.txt"

$Total = $Computers.Count
$Online = 0
$Offline = 0
$Users = 0
$32Bit = 0

ForEach ($Computer in $Computers)
{
    If (Test-Connection $Computer -Quiet -BufferSize 16 -Ea 0 -Count 1)
    {
        Try
        {
            $User = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer
            $Bit = $User.SystemType
            If ($Bit -like "x86*")
            {
                $32Bit += 1
            }
            If ($User.UserName -ne $null)
            {
                $EDI = $User.UserName.TrimStart("AREA52\")
                $UserInfo = (Get-ADUser "$EDI" -Properties DisplayName).DisplayName
                $Users +=1
            }
            Else
            {
                $UserInfo = "No user logged on"
            }
        }
        Catch
        {
            $UserInfo = "No access"
        }  
        
        Write-Host "$Computer ; online ; $Bit ; $UserInfo" -ForegroundColor Green
        $Online += 1
    }
    Else
    {
        Write-Host "$Computer ; offline" -ForegroundColor Red
        $Offline += 1
    }
}

$PercentG = [Math]::Round(($Online / $Total) * 100, 0)
$PercentB = [Math]::Round(($Offline / $Total) * 100, 0)
$PercentU = [Math]::Round(($Users / $Online) * 100, 0)
$Percent32 = [Math]::Round(($32Bit / $Online) * 100, 0)
$End = Get-Date
$TimeS = ($End - $Start).Seconds
$TimeM = ($End - $Start).Minutes

Write-Host
Write-Host "Online: $Online - $PercentG%"
Write-Host "Offline: $Offline - $PercentB%"
Write-Host "32 Bit: $32Bit - $Percent32%"
Write-Host "Users: $Users - $PercentU%"
Write-Host
Write-Host "Run Time: $TimeM min $TimeS sec"