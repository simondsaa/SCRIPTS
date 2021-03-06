    ####################################################
    # Change these values to the appropriate values in your environment

    $PrinterIP = "131.55.185.33"
    $PrinterPort = "9100"
    $PrinterPortName = "IP_" + $PrinterIP
    $DriverName = "HP Color LaserJet Flow MFP M577 PCL 6"
    $DriverPath = "C:\WINDOWS\system32\spool\drivers\x64\3\"
    $DriverInf = "C:\WINDOWS\system32\spool\drivers\x64\3\mxdwdrv.inf"
    $PrinterCaption = "HP Color LaserJet Flow MFP M577 PCL 6"
    ####################################################

    ### ComputerList Option 1 ###
    $ComputerList = get-content C:\Temp\machines.txt

    ### ComputerList Option 2 ###
    # $ComputerList = @()
    # Import-Csv "C:\Temp\ComputersThatNeedPrinters.csv" | `
    # % {$ComputerList += $_.Computer}

    Function CreatePrinterPort {
    param ($PrinterIP, $PrinterPort, $PrinterPortName, $ComputerName)
    $wmi = [wmiclass]"\\$ComputerName\root\cimv2:win32_tcpipPrinterPort"
    $wmi.psbase.scope.options.enablePrivileges = $true
    $Port = $wmi.createInstance()
    $Port.name = $PrinterPortName
    $Port.hostAddress = $PrinterIP
    $Port.portNumber = $PrinterPort
    $Port.SNMPEnabled = $false
    $Port.Protocol = 1
    $Port.put()
    }

    Function CreatePrinter {
    param ($PrinterCaption, $PrinterPortName, $DriverName, $ComputerName)
    $wmi = ([WMIClass]"\\$ComputerName\Root\cimv2:Win32_Printer")
    $Printer = $wmi.CreateInstance()
    $Printer.Caption = $PrinterCaption
    $Printer.DriverName = $DriverName
    $Printer.PortName = $PrinterPortName
    $Printer.DeviceID = $PrinterCaption
    $Printer.Put()
    }

    foreach ($computer in $ComputerList) {
     CreatePrinterPort -PrinterIP $PrinterIP -PrinterPort $PrinterPort `
     -PrinterPortName $PrinterPortName -ComputerName $computer
     InstallPrinterDriver -DriverName $DriverName -DriverPath `
     $DriverPath -DriverInf $DriverInf -ComputerName $computer
     CreatePrinter -PrinterPortName $PrinterPortName -DriverName `
     $DriverName -PrinterCaption $PrinterCaption -ComputerName $computer
    }
    ####################################################