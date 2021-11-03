Visual Basic Script

Copy
    '======================================================================================= 
    ' 
    ' NAME: WMIScan.vbs 
    ' 
    ' AUTHOR: Dougeby , Microsoft Corporation 
    ' DATE  : 10/24/2013 (Revised for System Center 2012 Configuration Manager by Rob Stack) 
    ' 
    ' COMMENT: Script to scan Configuration Manager WMI classes. 
    ' 
    '======================================================================================= 

    Dim SearchChar 
    Dim TotChar 
    Dim RightChar 
    Dim ClassName 
    Dim Computer 
    Dim strComputer 
    Dim strUser 
    Dim strPassword 
    Dim strSiteCode 
    Dim strNameSpace 
    Dim strFolder 
    Dim strFile 
    Dim strLogFile 
    Dim strFullFile 
    Dim strFullLogFile 
    Dim isError 

    Const ForWriting = 2 
    Const ForAppending = 8 
    Const adOpenStatic = 3 
    Const adLockOptimistic = 3 
    Const adUseClient = 3 

    set colNamedArguments=wscript.Arguments.Named 
    If colNamedArguments.Exists("Sitecode") Then 
      strSiteCode = colNamedArguments.Item("Sitecode") 
    Else 
      WScript.Echo "Invalid Command Line Arguments" & vbCrLf & _ 
        vbCrLf & "Usage: WMIScan.vbs /Sitecode:<sitecode> " & _ 
        "/Computer:<computername>" & vbCrLf & vbCrLf & _ 
        "Example1: WMIScan.vbs /Sitecode:PS1" & vbCrLf & _ 
        "Example2: WMIScan.vbs /Sitecode:PS1 /Computer:Computer1" 
      WScript.Quit(1) 
    End If 
    If colNamedArguments.Exists("Computer") Then 
      strComputer = colNamedArguments.Item("Computer") 
    Else strComputer = "." 
    End If 

    'Define the values for files and folders. 
    strFolder = "c:\WMIScan" 
    strFile = "WMIScan.txt" 
    strLogFile = "WMIScan.log" 
    strFullFile = strFolder & "\" & strFile 
    strFullLogFile = strFolder & "\" & strLogFile 
    isError = 0 

    'List of Configuration Manager namespaces are put into an array. 
    arrNameSpaces = Array("root\ccm","root\ccm\CCMPasswordSettings","root\ccm\CIModels",_
    "root\ccm\CIStateStore","root\ccm\CIStore","root\ccm\CITasks",_
    "root\ccm\ClientSDK","root\ccm\ContentTransferManager","root\ccm\DataTransferService",_
    "root\ccm\dcm","root\ccm\DCMAgent","root\ccm\evaltest",_
    "root\ccm\Events","root\ccm\InvAgt","root\ccm\LocationServices",_
    "root\ccm\Messaging","root\ccm\NetworkConfig","root\ccm\PeerDPAgent",_
    "root\ccm\Policy","root\ccm\PowerManagementAgent","root\ccm\RebootManagement",_
    "root\ccm\ScanAgent","root\ccm\Scheduler","root\ccm\SMSNapAgent",_
    "root\ccm\SoftMgmtAgent","root\ccm\SoftwareMeteringAgent","root\ccm\SoftwareUpdates",_
    "root\ccm\StateMsg","root\ccm\VulnerabilityAssessment","root\ccm\XmlStore",_
    "root\cimv2\sms","root\smsdm","root\sms",_
    "root\sms\site_"& strSiteCode)

    'Creates the folder and files for the scan output and log file. 
    Set objFSO = CreateObject("Scripting.FileSystemObject") 

    'Does strFolder Folder exist? If not, it's created. 
    If Not objFSO.FolderExists(strFolder) then 
      Set objFolder = objFSO.CreateFolder(strFolder) 
    End If 

    'Creates the WMIScan.txt and WMIScan.log files. 
    Set objFile = objFSO.CreateTextFile(strFullFile) 
    Set objLogFile = objFSO.CreateTextFile(strFullLogFile) 
    objFile.close 
    objLogFile.close 

    'Opens the WMIScan.log file in write mode. 
    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    Set objLogFile = objFSO.OpenTextFile(strFullLogFile, ForWriting) 
    objLogFile.WriteLine "********************************************" 
    objLogFile.WriteLine " WMIScan Tool Executed - " & Now() 
    objLogFile.WriteLine "********************************************" 

    'Opens the WMIScan.txt file in write mode. 
    Set objFile = objFSO.OpenTextFile(strFullFile, ForWriting) 
    objLogFile.WriteLine "--------------------------------------------" 
    Computer = strComputer 
    If Computer = "." Then Computer = "Local System" 
    objLogFile.WriteLine " Scanning WMI Namespaces On " & Computer 
    objLogFile.WriteLine "--------------------------------------------" 

    WScript.echo "Starting WMI scan on " & Computer 

    'Create a collection of Namespaces from the array, and 
    ' call the EnumNameSpaces subroutine to do the scan. 
    For Each strNameSpace In arrNameSpaces 
       Call EnumNameSpaces(strNameSpace, strComputer) 
    Next 
    objLogFile.WriteLine "---------------------------------------------" 
    objLogFile.WriteLine " Done scanning WMI Namespaces on " & Computer 
    objLogFile.WriteLine "---------------------------------------------" 

    'Close the WMISscan.txt file. 
    objFile.close 

    If isError = 1 Then 
      WScript.Echo "WMI Scan has Completed with Errors!" & vbCrLf & _ 
      "Check the " & strLogFile & " file for more details." & vbCrLf & _ 
      vbCrLf & strFile & " & " & strLogFile & " have been written to "_ 
      & strFolder & "." 
    Else 
      WScript.Echo "WMI Scan has Completed without any Errors!" & _ 
      vbCrLf & vbCrLf & strFile & " & " & strLogFile & _ 
      " have been written to " & strFolder & "." 
    End If   

    '*************************************************************** 
    '***   Subroutine to do the classes scan on the namespace.   *** 
    '*************************************************************** 
    Sub EnumNameSpaces(strNameSpace, strComputer) 
      Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator") 
      On Error Resume next 
      Set objSWbemServices= objSWbemLocator.ConnectServer (strComputer,_ 
        strNameSpace) 
      objLogFile.Write "Connecting to the \\" & strComputer & "\" &_ 
        strNameSpace & " WMI NameSpace...." 
      If Err.number = 0 Then  
        objLogFile.WriteLine "Success!!" 
        objLogFile.Write "  Scanning for Classes in "&strNameSpace _ 
          & "..." 

        'Create a collection of all the subclasses of the namespace. 
        Set colClasses = objSWbemServices.SubclassesOf() 

        'Scan all WMI classes, and write them to the scan1.txt file. 
        objFile.WriteBlanklines(1) 
        objFile.WriteLine "\\" & strComputer & "\" & strNameSpace 

        For Each objClass In colClasses 
          SearchChar = instr(objClass.Path_.Path, ":") 
          TotChar = len(objClass.Path_.Path) 
          RightChar = TotChar - SearchChar 
          ClassName = right(objClass.Path_.Path,RightChar) 
          objFile.WriteLine "   " & ClassName 
        Next 
        objLogFile.WriteLine "Success!!" 
      ElseIf Err.Number = -2147024891 Then 
        objLogFile.WriteLine "Error " & Err.Number & _ 
          "! Connection to "& strComputer & " Failed!" 
        isError = 1 
      Elseif Err.Number = -2147217394 Then 
        objLogFile.WriteLine "Error " & Err.Number & "!! Namespace "&_ 
          strNameSpace & " NOT Found!!" 
        isError = 1   
      Else 
        objLogFile.WriteLine "Error " & Err.Number & "!!" 
      isError = 1 
      End If 

    End Sub