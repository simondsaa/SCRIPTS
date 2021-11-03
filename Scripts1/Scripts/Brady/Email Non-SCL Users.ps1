$SNonSCL = $null
$ANonSCL = $null

$starttimer = Get-Date

# Save path for Excel sheet
$Path = "C:\Users\1392134782A\Desktop\Non-SCL Accounts.xlsx"
If (Test-Path $Path)
{
    Remove-Item -Path $Path -Force
}

# Variables for Email
$NetAdmin = "325CS.SCOO.NetAdmin@us.af.mil"
$Notice = "*This is an automatically generated email from PowerShell*"
$Message = "Attached is a listing of all the Non-SCL accounts in the Tyndall AFB OUs."

$starttimer = Get-Date

# Opening Excel
$a = New-Object -comobject Excel.Application
$a.visible = $True

$b = $a.Workbooks.Add(1)
$c = $b.Worksheets.Item(1)

# Column titles
$c.Cells.Item(1,1) = "SCL Required"
$c.Cells.Item(1,2) = "Display Name"
$c.Cells.Item(1,3) = "Account Name"
$c.Cells.Item(1,4) = "Logon Name"
$c.Cells.Item(1,5) = "Organization"
$c.Cells.Item(1,6) = "Description"
$c.Cells.Item(1,7) = "Account State"
$c.Cells.Item(1,8) = "Account Directory"
$c.Cells.Item(1,9) = "Total Numbers"

# Title cell formatting
$d = $c.UsedRange
$d.Interior.ColorIndex = 19
$d.Font.ColorIndex = 11
$d.Font.Bold = $True

$intRow = 2

# Pulls users from OUs that are SCL Exempt
$Users = Get-ADUser -Filter * -SearchBase "OU=Tyndall AFB,OU=AFCONUSEAST,OU=Bases,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" -Properties DisplayName, canonicalName, Organization, Description, SmartcardLogonRequired | Where {$_.SmartcardLogonRequired -eq $false}
$SUsers = Get-ADUser -Filter * -SearchBase "OU=Tyndall AFB,OU=Service Accounts,OU=Administration,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" -Properties DisplayName, canonicalName, Organization, Description, SmartcardLogonRequired
$AUsers = Get-ADUser -Filter * -SearchBase "OU=Tyndall AFB,OU=Administrative Accounts,OU=Administration,DC=AREA52,DC=AFNOAPPS,DC=USAF,DC=MIL" -Properties DisplayName, canonicalName, Organization, Description, SmartcardLogonRequired

# Users Accounts
ForEach ($User in $Users)
{
    If ($User.Enabled -eq $true)
    {
        $State = "Enabled"
    }
    Else
    {
        $State = "Disabled"
    }
    
    $c.Cells.Item($intRow,1) = "FALSE"        
    $c.Cells.Item($intRow,2) = $User.DisplayName
    $c.Cells.Item($intRow,3) = $User.Name
    $c.Cells.Item($intRow,4) = $User.UserPrincipalName
    $c.Cells.Item($intRow,5) = $User.Organization
    $c.Cells.Item($intRow,6) = $User.Description
    $c.Cells.Item($intRow,7) = $State
    $c.Cells.Item($intRow,8) = $User.CanonicalName

    $intRow = $intRow + 1
}

# Service Accounts
ForEach ($SUser in $SUsers)
{
    If ($SUser.Enabled -eq $true)
    {
        $State = "Enabled"
    }
    Else
    {
        $State = "Disabled"
    }

    If ($SUser.SmartcardLogonRequired -eq $false)
    {
        $SNonSCL += 1
        $c.Cells.Item($intRow,1) = "FALSE"
    }
    Else
    {
        $c.Cells.Item($intRow,1) = "TRUE"
    }
    
    $c.Cells.Item($intRow,2) = $SUser.DisplayName
    $c.Cells.Item($intRow,3) = $SUser.Name
    $c.Cells.Item($intRow,4) = $SUser.UserPrincipalName
    $c.Cells.Item($intRow,5) = $SUser.Organization
    $c.Cells.Item($intRow,6) = $SUser.Description
    $c.Cells.Item($intRow,7) = $State
    $c.Cells.Item($intRow,8) = $SUser.CanonicalName
    
    $intRow = $intRow + 1
}

# Admin Accounts
ForEach ($AUser in $AUsers)
{
    If ($AUser.Enabled -eq $true)
    {
        $State = "Enabled"
    }
    Else
    {
        $State = "Disabled"
    }
    
    If ($AUser.SmartcardLogonRequired -eq $false)
    {
        $ANonSCL += 1
        $c.Cells.Item($intRow,1) = "FALSE"
    }
    Else
    {
        $c.Cells.Item($intRow,1) = "TRUE"
    }

    $c.Cells.Item($intRow,2) = $AUser.DisplayName
    $c.Cells.Item($intRow,3) = $AUser.Name
    $c.Cells.Item($intRow,4) = $AUser.UserPrincipalName
    $c.Cells.Item($intRow,5) = $AUser.Organization
    $c.Cells.Item($intRow,6) = $AUser.Description
    $c.Cells.Item($intRow,7) = $State
    $c.Cells.Item($intRow,8) = $AUser.CanonicalName
    
    $intRow = $intRow + 1
}

$c.Cells.Item(2,9) = "Total Non-SCL Users"
$c.Cells.Item(3,9) = "Total Non-SCL Service"
$c.Cells.Item(4,9) = "Total Non-SCL Admin"
$c.Cells.Item(5,9) = "Total Admin Accounts"

$c.Cells.Item(2,10) = $Users.Count
$c.Cells.Item(3,10) = $SNonSCL
$c.Cells.Item(4,10) = $ANonSCL
$c.Cells.Item(5,10) = $AUsers.Count

# Excel formatting to close and save
$d.EntireColumn.AutoFit()

$b.SaveAs($Path)
$b.Close()

$a.Quit()

# Body for email
$Body = "Lt Mayers,

$Message

$Notice

V/R

Network Operations
325 Communications Squadron
Tyndall AFB, FL 32403
COMM 850-283-8230
DSN 523-8230"

# Sending the mail message with email options
#Send-MailMessage -From $NetAdmin -To paul.mayers.2@us.af.mil -Cc michael.jones.218@us.af.mil -Bcc $NetAdmin -Priority High -Attachments $Path -Subject "Non-SCL Accounts" -Body $Body -SmtpServer wrightpatterson.oa.us.af.mil

$stoptimer = Get-Date
"Execution Time: {0} Minutes" -f [Math]::Round(($stoptimer - $starttimer).TotalMinutes , 2)