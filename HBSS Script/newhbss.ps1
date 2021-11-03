function Install-McAfee {
    [CmdletBinding()]
    param(
        [Parameter()]
        [System.Management.Automation.Runspaces.PSSession] $session,
        [string] $FilePath
    )

    process {
    try {
        Copy-Item $FilePath -ToSession $session -Destination 'C:\FramePkg.exe' -ErrorAction Stop
        Invoke-Command -Session $session -ScriptBlock {C:\FramePkg.exe  '/install=agent' /s | Out-Null -ErrorAction Stop }
        } catch {
       
        }
    }
} 

function Uninstall-McAfee {
    [CmdletBinding()]
    param(
        [Parameter()]
        [System.Management.Automation.Runspaces.PSSession] $session
    )

    process {
    try {
        Invoke-Command -Session $session -ScriptBlock {
        Set-Location 'C:\Program Files\McAfee\Agent\x86\'
        ./FrmInst.exe /forceuninstall /s | Out-Null -ErrorAction Stop
        } 
    } catch {
      
        }
    }
} 

function Get-Version {
    [CmdletBinding()]
    param([Parameter()]
        [System.Management.Automation.Runspaces.PSSession] $session,
        [int] $LastetMcAfeeVersion)

    process {
     try {
        [string]$version = Invoke-Command -Session $session -ScriptBlock { (Get-ItemProperty 'C:\Program Files\McAfee\Agent\cmdagent.exe').VersionInfo.FileVersion } -ErrorAction Stop
        $version = $version -replace "\.", ""
        if([int]$version -ge $LastetMcAfeeVersion) {
        return $true } else { return $false }
    } catch { 
        return $false 
        }
    }
} 

function Check-PSSession {
    [CmdletBinding()]
    param(
        [Parameter()]
        [pscustomobject] $computername
    )

    process {

    try {
        New-PSSession -ComputerName $computername -ErrorAction Stop
        return $true
    } 
    catch 
    {
        return $false
    }

  }

} 

function Export-Result {
    [CmdletBinding()] 
    param([Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.PSSession] $session,
        [pscustomobject] $ComputerName,
        [string] $FilePath
    )
    
    process {
    try { 
    $validRegPath = Invoke-Command -Session $session -ScriptBlock { Test-Path -Path 'HKLM:\SOFTWARE\WOW6432Node\Network Associates\ePolicy Orchestrator\agent' }  
    if ($validRegPath) { 
    Invoke-Command -Session $session -ScriptBlock { Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\Network Associates\ePolicy Orchestrator\agent' | Select-Object -Property @{Name = 'Computer Name'; Expression = {[System.Net.Dns]::GetHostName()}}, @{Name = 'IPv4 Address'; Expression = {get-netipaddress | Select-Object -ExpandProperty ipaddress | Where-Object {$_ -like '131.*'}}}, @{Name = 'McAfee Version'; Expression = {(Get-ItemProperty 'C:\Program Files\McAfee\Agent\cmdagent.exe').VersionInfo.FileVersion}}, AgentGUID, @{Name = 'LastASCTime'; Expression = {[timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').addSeconds([int]$_.LastASCTime))}}, @{Name = 'LastPolicyUpdateTime'; Expression = {[timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').addSeconds([int]$_.LastPolicyUpdateTime))}}, ePOServerList } | Export-Csv -Path $FilePath -NoTypeInformation -Append
    } else {
        $ComputerName | Select-Object @{Name = 'Computer Name'; Expression = {$ComputerName}} | Export-Csv -Path $FilePath -NoTypeInformation -Append 
    }
    } catch {

        }
    }
}

function Remove-FramePkg {

    [CmdletBinding()] 
    param([Parameter()]
        [System.Management.Automation.Runspaces.PSSession] $session
    )

    process {

    try {

    $validPath = Invoke-Command -Session $session -ScriptBlock { Test-Path -Path 'C:\FramePkg.exe' }

    if($validPath) {
    Remove-Item -Path C:\FramePkg.exe -
    }

    } catch {

    }

    }

}

<#
 Examples
 $RogueFilePath = C:\Users\1555519520a.adw\Desktop\pc.csv
 $SucessfulFilePath = 'C:\success.csv'
 $UnsucessfulFilePath = 'C:\unsuccess.csv'
#>

$RogueFilePath = Read-Host "`nPlease enter file path of Rogue List"
$FramePackageFilePath = Read-Host "`nPlease enter file path of FramePkg.exe"
$SucessfulFilePath = Read-Host "`nPlease enter a .csv file path for the successful output"
$UnsucessfulFilePath = Read-Host "`nPlease enter a .csv file path for the unsuccessful output"
$LastestMcAfeeVersion = Read-Host "`nPlease enter the lastest McAfee version without the decimals, example: 5.6.2.209 -eq 562209"

Import-Csv -Path $RogueFilePath | ForEach-Object {

$ComputerName = $_.ComputerName

write-host ""
write-host "Working on $ComputerName" -Foreground yellow

$canRemote = Check-PSSession -computername $ComputerName

if($canRemote) {
    $session = New-PSSession -ComputerName $ComputerName
    $validMcAfee = Get-Version -session $session -LastetMcAfeeVersion $LastestMcAfeeVersion
    $validPath = Invoke-Command -Session $session -ScriptBlock { Test-Path -Path "C:\Program Files\McAfee\Agent\x86\FrmInst.exe" } 
    if ($validMcAfee) {
    Export-Result -session $session -FilePath $SucessfulFilePath
    write-host ""
    write-host "$ComputerName has been updated" -Foreground Green
    Remove-FramePkg -session $session
    } else {
    if ($validPath) {
    write-host ""
    write-host "Uninstalling the Frame Package on $ComputerName" -Foreground Blue
    Uninstall-McAfee -session $session
    write-host ""
    write-host "Installing the lastest Frame Package on $ComputerName" -Foreground Blue
    Install-McAfee -FilePath $FramePackageFilePath -session $session
    Export-Result -session $session -ComputerName $ComputerName -FilePath $SucessfulFilePath
    write-host ""
    write-host "$ComputerName has been updated" -Foreground Green
    Remove-FramePkg -session $session
    } else {
    write-host ""
    write-host "Installing the lastest Frame Package on $ComputerName" -Foreground Blue
    Install-McAfee -FilePath $FramePackageFilePath -session $session
    Export-Result -session $session -ComputerName $ComputerName -FilePath $SucessfulFilePath
    write-host ""
    write-host "$ComputerName has been updated" -Foreground Green
    Remove-FramePkg -session $session
        }
    }
}
else 
{
    $ComputerName | Select-Object @{Name = 'Computer Name'; Expression = {$ComputerName}} | Export-Csv -Path $UnsucessfulFilePath -NoTypeInformation -Append 
    write-host ""
    write-host "$ComputerName is not online" -Foreground Red
}
    
}


 
