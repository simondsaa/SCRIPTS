# Specify the Computers you want to query:
$Computers = Get-Content "C:\Users\1252862141.adm\Desktop\Scripts\PC_Info.txt"
ForEach ($Computer in $Computers){
    Write-Host
    Write-Host "BEGIN REPORT-------------------------------------------------------"
    $Comp = Get-WmiObject Win32_ComputerSystem -comp $Computer -ErrorAction SilentlyContinue
    Write-Host -ForeGroundColor White "Computer Name:    " $Comp.Name
    Write-Host -ForeGroundColor White "User Logged On:   " $Comp.UserName   
        $OS = Get-Wmiobject Win32_OperatingSystem -comp $Computer -ErrorAction SilentlyContinue
        Write-Host -ForeGroundColor Yellow "Operating System: " $OS.Caption"SP" $OS.ServicePackMajorVersion
        Write-Host -ForeGroundColor Yellow "Installed On:     " $OS.ConvertToDateTime($OS.InstallDate)  
            $NIC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $Computer -ErrorAction SilentlyContinue
            Write-Host -ForeGroundColor White "IP Address:       " $NIC.IPAddress
            Write-Host -ForeGroundColor White "MAC Address:      " $NIC.MACAddress
    Write-Host "END OF REPORT------------------------------------------------------"
    Write-Host}