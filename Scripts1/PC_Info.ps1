# Specify the Computers you want to query:
$Computers = Get-Content "C:\Temp\test.txt"
ForEach ($Computer in $Computers){
    Write-Host
    Write-Host "BEGIN REPORT-------------------------------------------------------"
    $Comp = Get-WmiObject Win32_ComputerSystem -comp $Computer -ErrorAction SilentlyContinue
    Write-Host -ForeGroundColor Yellow "Computer Name:    " $Comp.Name
    Write-Host -ForeGroundColor Yellow "User Logged On:   " $Comp.UserName
    Write-Host -ForeGroundColor Yellow "Model             " $Comp.Model   
        $OS = Get-Wmiobject Win32_OperatingSystem -comp $Computer -ErrorAction SilentlyContinue
        Write-Host -ForeGroundColor White "Operating System: " $OS.Caption"SP" $OS.ServicePackMajorVersion
        Write-Host -ForeGroundColor White "Installed On:     " $OS.ConvertToDateTime($OS.InstallDate)  
            $NIC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -filter "IPEnabled='True'" -cn $Computer -ErrorAction SilentlyContinue
            Write-Host -ForeGroundColor White "IP Address:       " $NIC.IPAddress
            Write-Host -ForeGroundColor White "MAC Address:      " $NIC.MACAddress
    Write-Host "END OF REPORT------------------------------------------------------"
    Write-Host}