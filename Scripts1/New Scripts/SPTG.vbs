' Logon6.vbs
' VBScript logon script program.
'
' ----------------------------------------------------------------------
' Copyright (c) 2004-2010 Richard L. Mueller
' Hilltop Lab web site - http://www.rlmueller.net
' Version 1.0 - March 28, 2004
' Version 1.1 - July 30, 2007 - Escape any "/" characters in DN's.
' Version 1.2 - November 6, 2010 - No need to set objects to Nothing.
'
' You have a royalty-free right to use, modify, reproduce, and
' distribute this script file in any way you find useful, provided that
' you agree that the copyright owner above has no warranty, obligations,
' or liability for such use.

Option Explicit

Dim objRootDSE, objTrans, strNetBIOSDomain, objNetwork, strNTName
Dim strUserDN, strComputerDN, objGroupList, objUser, strDNSDomain
Dim strComputer, objComputer
Dim strHomeDrive, strHomeShare
Dim adoCommand, adoConnection, strBase, strAttributes
Dim objShell		'Creates Windows Shell Object
Dim oShell
Dim StrCmd
Dim bForce, bUpdateProfile
Dim objFSO

' Constants for the NameTranslate object.
Const ADS_NAME_INITTYPE_GC = 3
Const ADS_NAME_TYPE_NT4 = 3
Const ADS_NAME_TYPE_1779 = 1

Set objNetwork = CreateObject("Wscript.Network")

' Loop required for Win9x clients during logon.
strNTName = ""
On Error Resume Next
Do While strNTName = ""
    strNTName = objNetwork.UserName
    Err.Clear
    If (Wscript.Version > 5) Then
        Wscript.Sleep 100
    End If
Loop
On Error GoTo 0

' Determine DNS domain name from RootDSE object.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

' Use the NameTranslate object to find the NetBIOS domain name from the
' DNS domain name.
Set objTrans = CreateObject("NameTranslate")
objTrans.Init ADS_NAME_INITTYPE_GC, ""
objTrans.Set ADS_NAME_TYPE_1779, strDNSDomain
strNetBIOSDomain = objTrans.Get(ADS_NAME_TYPE_NT4)
' Remove trailing backslash.
strNetBIOSDomain = Left(strNetBIOSDomain, Len(strNetBIOSDomain) - 1)

' Use the NameTranslate object to convert the NT user name to the
' Distinguished Name required for the LDAP provider.
objTrans.Set ADS_NAME_TYPE_NT4, strNetBIOSDomain & "\" & strNTName
strUserDN = objTrans.Get(ADS_NAME_TYPE_1779)
' Escape any forward slash characters, "/", with the backslash
' escape character. All other characters that should be escaped are.
strUserDN = Replace(strUserDN, "/", "\/")

' Bind to the user object in Active Directory with the LDAP provider.
Set objUser = GetObject("LDAP://" & strUserDN)

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Map a network drive if the user is a member of the group.

'If (IsMember(objUser, "325 CS SCOO ALL") = True) Then
'    On Error Resume Next
'Set objShell = CreateObject("Wscript.Shell")
'objShell.Run "\\xlwu-fs-05pv\Tyndall_PUBLIC\NCC Admin\BGINFO\Bginfo.bgi"
'End If

If (IsMember(objUser, "!SG") = True) Then
    	On Error Resume Next
    objNetwork.MapNetworkDrive "T:", "\\xlwu-fs-05pv\Tyndall_PUBLIC"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "T:", True, True
        objNetwork.MapNetworkDrive "T:", "\\xlwu-fs-05pv\Tyndall_PUBLIC"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 CS_ALL") = True) Then
    	On Error Resume Next
    objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CS"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "S:", True, True
        objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CS"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 MXG_STAFF") = True) Then
    	On Error Resume Next
    objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\325 MXG STAFF"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "S:", True, True
        objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\325 MXG STAFF"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 AMXS_ALL") = True) Then
    	On Error Resume Next
    objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\325 AMXS"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "S:", True, True
        objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\325 AMXS"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_43 AMU_ALL") = True) Then
    	On Error Resume Next
    objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\325 AMXS\43 AMU"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "S:", True, True
        objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\325 AMXS\43 AMU"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_95 AMU_ALL") = True) Then
    	On Error Resume Next
    objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\95 AMU"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "S:", True, True
        objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\95 AMU"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 MXS_ALL") = True) Then
    	On Error Resume Next
    objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\325 MXS"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "S:", True, True
        objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\325 MXS"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_372 TRS_ALL") = True) Then
    	On Error Resume Next
    objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\372 TRS"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "S:", True, True
        objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\372 TRS"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_LOCKHEED_ALL") = True) Then
    	On Error Resume Next
    objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\Lockheed"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "S:", True, True
        objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MXG\Lockheed"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 SFS_ALL") = True) Then
    On Error Resume Next
    objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 SFS"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "S:", True, True
        objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 SFS"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "325 SFS All") = True) Then
    On Error Resume Next
    objNetwork.MapNetworkDrive "N:", "\\xlwu-fs-05pv\Tyndall_PUBLIC\Applications\FEDLOG"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "N:", True, True
        objNetwork.MapNetworkDrive "N:", "\\xlwu-fs-05pv\Tyndall_PUBLIC\Applications\FEDLOG"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 CES_All") = True) Then
    On Error Resume Next
    objNetwork.RemoveNetworkDrive "X:"
    objNetwork.MapNetworkDrive "X:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CES"
 If (Err.Number <> 0) Then
        On Error GoTo 0
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 CES_CEN") = True) Then
    On Error Resume Next
    objNetwork.RemoveNetworkDrive "X:"
    objNetwork.MapNetworkDrive "X:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CES\CEN"
    objNetwork.RemoveNetworkDrive "M:"
    objNetwork.MapNetworkDrive "M:", "\\TYNAP001P1\maint"
    objNetwork.RemoveNetworkDrive "P:"
    objNetwork.MapNetworkDrive "P:", "\\TYNAP001P1\DWG103"
    objNetwork.RemoveNetworkDrive "L:"
    objNetwork.MapNetworkDrive "L:", "\\TYNAP001P1\PROGENL"
    objNetwork.RemoveNetworkDrive "V:"
    objNetwork.MapNetworkDrive "V:", "\\TYNAP001P1\EVAULT"
    If (Err.Number <> 0) Then
        On Error GoTo 0
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 CES_CED") = True) Then
    On Error Resume Next
    objNetwork.MapNetworkDrive "X:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CES\CED"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "X:", True, True
        objNetwork.MapNetworkDrive "X:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CES\CED"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 CES_CEF") = True) Then
    On Error Resume Next
    objNetwork.RemoveNetworkDrive "X:"
    objNetwork.MapNetworkDrive "X:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CES\CEF"
 If (Err.Number <> 0) Then
        On Error GoTo 0
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 CES_CEO") = True) Then
    On Error Resume Next
    objNetwork.RemoveNetworkDrive "X:"
    objNetwork.MapNetworkDrive "X:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CES\CEO"
 If (Err.Number <> 0) Then
        On Error GoTo 0
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "325CES_CEO") = True) Then
    On Error Resume Next
    objNetwork.RemoveNetworkDrive "V:"
    objNetwork.MapNetworkDrive "V:", "\\TYNAP001P1\EVAULT"
    objNetwork.RemoveNetworkDrive "M:"
    objNetwork.MapNetworkDrive "M:", "\\TYNAP001P1\maint"
    objNetwork.RemoveNetworkDrive "P:"
    objNetwork.MapNetworkDrive "P:", "\\TYNAP001P1\DWG103"
    objNetwork.RemoveNetworkDrive "L:"
    objNetwork.MapNetworkDrive "L:", "\\TYNAP001P1\PROGENL"
    If (Err.Number <> 0) Then
        On Error GoTo 0
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 CES_CC") = True) Then
    On Error Resume Next
    objNetwork.RemoveNetworkDrive "X:"
    objNetwork.MapNetworkDrive "X:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CES\CC"
 If (Err.Number <> 0) Then
        On Error GoTo 0
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 CES_CEX") = True) Then
    On Error Resume Next
    objNetwork.RemoveNetworkDrive "X:"
    objNetwork.MapNetworkDrive "X:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CES\CEX"
 If (Err.Number <> 0) Then
        On Error GoTo 0
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 CES_BOS") = True) Then
    On Error Resume Next
    objNetwork.RemoveNetworkDrive "X:"
    objNetwork.MapNetworkDrive "X:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CES\BOS"
 If (Err.Number <> 0) Then
        On Error GoTo 0
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 FW_STAFF") = True) Then
    On Error Resume Next
    objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_FW\325 FW Staff"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "S:", True, True
        objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_FW\325 FW Staff"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "325CONS_All") = True) Then
    On Error Resume Next
    objNetwork.RemoveNetworkDrive "O:"
    objNetwork.MapNetworkDrive "O:", "\\tynfs10\ABSS"
    objNetwork.RemoveNetworkDrive "W:"
    objNetwork.MapNetworkDrive "W:", "\\afcis-apps\FTP"
    If (Err.Number <> 0) Then
        On Error GoTo 0
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "325CONS_SA") = True) Then
    On Error Resume Next
    objNetwork.MapNetworkDrive "M:", "\\131.55.130.3\SPS-I"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "M:", True, True
        objNetwork.MapNetworkDrive "M:", "\\131.55.130.3\SPS-I"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 LRS_ALL") = True) Then
    On Error Resume Next
    objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 LRS"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "S:", True, True
        objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 LRS"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_325 CONS_ALL") = True) Then
    On Error Resume Next
    objNetwork.MapNetworkDrive "R:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CONS"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "R:", True, True
        objNetwork.MapNetworkDrive "R:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 CONS"
    End If
    On Error GoTo 0
End If


If (IsMember(objUser, "SUP_SUPPLY") = True) Then
    On Error Resume Next
    objNetwork.RemoveNetworkDrive "F:"
    objNetwork.MapNetworkDrive "F:", "\\TYNSUP02\FED LOG" 
    objNetwork.RemoveNetworkDrive "J:"
    objNetwork.MapNetworkDrive "J:", "\\FS-TYNSUP-03\THUD"
    objNetwork.RemoveNetworkDrive "P:"
    objNetwork.MapNetworkDrive "P:", "\\TYNSUPPLY\APPLICATIONS"
    objNetwork.RemoveNetworkDrive "S:"
    objNetwork.MapNetworkDrive "S:", "\\TYNSUPPLY\SHAREDFILES"
    objNetwork.RemoveNetworkDrive "V:"
    objNetwork.MapNetworkDrive "V:", "\\TYNSUP02\APPLICATIONS\AF\AFILSEPL"
    If (Err.Number <> 0) Then
        On Error GoTo 0
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "SUP_FUELS") = True) Then
    On Error Resume Next
    objNetwork.RemoveNetworkDrive "F:"
    objNetwork.MapNetworkDrive "F:", "\\TYNSUP02\FCCDATA"
    objNetwork.RemoveNetworkDrive "I:"
    objNetwork.MapNetworkDrive "I:", "\\TYNSUPPLY\FUELS FILES"
    objNetwork.RemoveNetworkDrive "S:"
    objNetwork.MapNetworkDrive "S:", "\\TYNSUPPLY\SHAREDFILES"
    If (Err.Number <> 0) Then
        On Error GoTo 0
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "SUP_TRANS") = True) Then
    On Error Resume Next
    objNetwork.RemoveNetworkDrive "S:"
    objNetwork.MapNetworkDrive "S:", "\\TYNSUPPLY\SHAREDFILES"
    If (Err.Number <> 0) Then
        On Error GoTo 0
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "DLS_44 FG_All") = True) Then
    On Error Resume Next
    objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-02pv\Tyndall_44_FG"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "S:", True, True
        objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-02pv\Tyndall_44_FG"
    End If
    On Error GoTo 0
End If

If (IsMember(objUser, "USG_325 FSS_FSS_ALL") = True) Then
    On Error Resume Next
    objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 FSS"
    If (Err.Number <> 0) Then
        On Error GoTo 0
        objNetwork.RemoveNetworkDrive "S:", True, True
        objNetwork.MapNetworkDrive "S:", "\\xlwu-fs-04pv\Tyndall_325_MSG\325 FSS"
    End If
    On Error GoTo 0
End If

' Use the NameTranslate object to convert the NT name of the computer to
' the Distinguished name required for the LDAP provider. Computer names
' must end with "$".
strComputer = objNetwork.computerName
objTrans.Set ADS_NAME_TYPE_NT4, strNetBIOSDomain _
    & "\" & strComputer & "$"
strComputerDN = objTrans.Get(ADS_NAME_TYPE_1779)
' Escape any forward slash characters, "/", with the backslash
' escape character. All other characters that should be escaped are.
strComputerDN = Replace(strComputerDN, "/", "\/")

' Bind to the computer object in Active Directory with the LDAP
' provider.
Set objComputer = GetObject("LDAP://" & strComputerDN)

' Add a printer connection if the computer is a member of the group.
If (IsMember(objComputer, "Room 231") = True) Then
    objNetwork.AddPrinterConnection "LPT1:", "\\PrintServer\Printer3"
End If

' Clean up.
If (IsObject(adoConnection) = True) Then
    adoConnection.Close
End If

Function IsMember(ByVal objADObject, ByVal strGroupNTName)
    ' Function to test for group membership.
    ' objADObject is a user or computer object.
    ' strGroupNTName is the NT name (sAMAccountName) of the group to test.
    ' objGroupList is a dictionary object, with global scope.
    ' Returns True if the user or computer is a member of the group.
    ' Subroutine LoadGroups is called once for each different objADObject.

    ' The first time IsMember is called, setup the dictionary object
    ' and objects required for ADO.
    If (IsEmpty(objGroupList) = True) Then
        Set objGroupList = CreateObject("Scripting.Dictionary")
        objGroupList.CompareMode = vbTextCompare

        Set adoCommand = CreateObject("ADODB.Command")
        Set adoConnection = CreateObject("ADODB.Connection")
        adoConnection.Provider = "ADsDSOObject"
        adoConnection.Open "Active Directory Provider"
        adoCommand.ActiveConnection = adoConnection

        Set objRootDSE = GetObject("LDAP://RootDSE")
        strDNSDomain = objRootDSE.Get("defaultNamingContext")

        adoCommand.Properties("Page Size") = 100
        adoCommand.Properties("Timeout") = 30
        adoCommand.Properties("Cache Results") = False

        ' Search entire domain.
        strBase = "<LDAP://" & strDNSDomain & ">"
        ' Retrieve NT name of each group.
        strAttributes = "sAMAccountName"

        ' Load group memberships for this user or computer into dictionary
        ' object.
        Call LoadGroups(objADObject)
    End If
    If (objGroupList.Exists(objADObject.sAMAccountName & "\") = False) Then
        ' Dictionary object established, but group memberships for this
        ' user or computer must be added.
        Call LoadGroups(objADObject)
    End If
    ' Return True if this user or computer is a member of the group.
    IsMember = objGroupList.Exists(objADObject.sAMAccountName & "\" _
        & strGroupNTName)
End Function

Sub LoadGroups(ByVal objADObject)
    ' Subroutine to populate dictionary object with group memberships.
    ' objGroupList is a dictionary object, with global scope. It keeps track
    ' of group memberships for each user or computer separately. ADO is used
    ' to retrieve the name of the group corresponding to each objectSid in
    ' the tokenGroup array. Based on an idea by Joe Kaplan.

    Dim arrbytGroups, k, strFilter, adoRecordset, strGroupName, strQuery

    ' Add user name to dictionary object, so LoadGroups need only be
    ' called once for each user or computer.
    objGroupList.Add objADObject.sAMAccountName & "\", True

    ' Retrieve tokenGroups array, a calculated attribute.
    objADObject.GetInfoEx Array("tokenGroups"), 0
    arrbytGroups = objADObject.Get("tokenGroups")

    ' Create a filter to search for groups with objectSid equal to each
    ' value in tokenGroups array.
    strFilter = "(|"
    If (TypeName(arrbytGroups) = "Byte()") Then
        ' tokenGroups has one entry.
        strFilter = strFilter & "(objectSid=" _
            & OctetToHexStr(arrbytGroups) & ")"
    ElseIf (UBound(arrbytGroups) > -1) Then
        ' TokenGroups is an array of two or more objectSid's.
        For k = 0 To UBound(arrbytGroups)
            strFilter = strFilter & "(objectSid=" _
                & OctetToHexStr(arrbytGroups(k)) & ")"
        Next
    Else
        ' tokenGroups has no objectSid's.
        Exit Sub
    End If
    strFilter = strFilter & ")"

    ' Use ADO to search for groups whose objectSid matches any of the
    ' tokenGroups values for this user or computer.
    strQuery = strBase & ";" & strFilter & ";" _
        & strAttributes & ";subtree"
    adoCommand.CommandText = strQuery
    Set adoRecordset = adoCommand.Execute

    ' Enumerate groups and add NT name to dictionary object.
    Do Until adoRecordset.EOF
        strGroupName = adoRecordset.Fields("sAMAccountName").Value
        objGroupList.Add objADObject.sAMAccountName & "\" _
            & strGroupName, True
        adoRecordset.MoveNext
    Loop
    adoRecordset.Close

End Sub

Function OctetToHexStr(ByVal arrbytOctet)
    ' Function to convert OctetString (byte array) to Hex string,
    ' with bytes delimited by \ for an ADO filter.

    Dim k
    OctetToHexStr = ""
    For k = 1 To Lenb(arrbytOctet)
        OctetToHexStr = OctetToHexStr & "\" _
            & Right("0" & Hex(Ascb(Midb(arrbytOctet, k, 1))), 2)
    Next
End Function

Set oShell = CreateObject("WScript.Shell")
strCmd = "powerpnt /s  \\xlwu-fs-05pv.area52.afnoapps.usaf.mil\Tyndall_PUBLIC\Base_Shares\Advertisements\Temp.pps"
oShell.Run(strCmd)
Set oShell = Nothing