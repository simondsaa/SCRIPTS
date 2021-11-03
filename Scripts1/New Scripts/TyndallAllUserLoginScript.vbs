' FileName:   TyndallAllUserLoginScript.vbs
' Usage:       for login by all Tyndall AFB to Area52 domain

'----------------------------------------------------------------------------------------
' Variable declaration/definition
'----------------------------------------------------------------------------------------
Option Explicit


'-----Variable Definitions:
'     blnIsMember   : true/false boolean
'     strUserID     : userid string
'     strComputer   : computer name string
'     strDomain     : domain name string
'     strGroupName  : AD group name string
'     strLetter     : drive letter string
'     strUNC        : UNC path to server share string
'     intWSHVersion : holds WSH version number integer
'     objWSH        : shell object
'     objFSO        : file system object
'     objNet        : net object
'     objNetwork    : network object
'     objUser       : user object
'     objGroup      : group object
' Initialize system objects

Dim objRootDSE, objTrans, strNetBIOSDomain, objNetwork, strNTName
Dim strUserDN, strComputerDN, objGroupList, objUser, strDNSDomain
Dim strComputer, objComputer
Dim strHomeDrive, strHomeShare
Dim adoCommand, adoConnection, strBase, strAttributes
Dim strDomain, strGroupName, strLetter, strUNC, objSysInfo
Dim blnIsMember, strUserID, strGroup, strUserCN
Dim intWSHVersion, objWSH, objFSO, objNet, objGroup, objShell, objUser2

set objWSH = createobject("WScript.shell")
set objFSO = createobject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("WScript.Network")

Function IsMemberOf(strGroupName)
  Set objNetwork = CreateObject("WScript.Network")
  strDomain = objNetwork.UserDomain
  strUserID = objNetwork.UserName
  blnIsMember = False
  Set objUser2 = GetObject("WinNT://" & strDomain & "/" & strUserID & ",user")
  For Each objGroup In objUser2.Groups
    If LCase(objGroup.Name) = LCase(strGroupName) Then
      blnIsMember = True
      Exit For
    End If
  Next
  IsMemberOf = blnIsMember
End Function

Sub MapDrive(strLetter, strUNC)
  Set objNet = WScript.CreateObject("WScript.Network")
  
  If objFSO.DriveExists(strLetter) Then
    objNet.RemoveNetworkDrive strLetter, True, True
    objNet.MapNetworkDrive strLetter, strUNC
  Else
    objNet.MapNetworkDrive strLetter, strUNC
  End If
End Sub

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

'Set objShell = CreateObject("Wscript.Shell")
'objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\PII.vbs //NOLOGO"

' Map a network drive if the user is a member of the group.

If (IsMember(objUser, "GLS_325 TRSS_JOIN ALL GTIMS USERS") = True) Then
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\AREA52_TYN_32bit.bat //NOLOGO"
End If

If (IsMember(objUser, "GLS_325 TRSS_JOIN ALL GTIMS USERS") = True) Then
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\AREA52_TYN_64bit.bat //NOLOGO"
End If

If (IsMember(objUser, "!479") = True) Then
    On Error Resume Next
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\479FTG.vbs //NOLOGO"
End If

If (IsMember(objUser, "!OG") = True) Then
    On Error Resume Next
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\OG.vbs //NOLOGO"
End If

If (IsMember(objUser, "!WEG") = True) Then
    On Error Resume Next
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\weg.vbs //NOLOGO"
End If

If (IsMember(objUser, "!AFNORTH") = True) Then
    On Error Resume Next
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\AFNORTH.vbs //NOLOGO"
End If

If (IsMember(objUser, "!RHS") = True) Then
    On Error Resume Next
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\sf.vbs //NOLOGO"
End If

If (IsMember(objUser, "!NCOA") = True) Then
    On Error Resume Next
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\NCOA.vbs //NOLOGO"
End If

If (IsMember(objUser, "!66TRS") = True) Then
    On Error Resume Next
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\66TRS.vbs //NOLOGO"
End If

If (IsMember(objUser, "!359TRS") = True) Then
    On Error Resume Next
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\359TRS.vbs //NOLOGO"
End If

If (IsMember(objUser, "!316TRS") = True) Then
    On Error Resume Next
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\316TRS.vbs //NOLOGO"
End If

If (IsMember(objUser, "AFCESA Users") = True) Then
    On Error Resume Next
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\AFCESA.vbs //NOLOGO"
End If

If (IsMember(objUser, "_G_325MDG") = True) Then
    On Error Resume Next
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\MDG.vbs //NOLOGO"
End If

If strGroup = ("Tyndall_MilMod") Then
objNetwork.MapNetworkDrive ("M:"), ("Tyndall_MilMod")
End If

If (IsMember(objUser, "AFRL Users") = True) Then
    On Error Resume Next
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "\\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\afrlmlq.bat //NOLOGO"
End If


Set objShell = CreateObject("Wscript.Shell")
 objShell.Run("powershell.exe -executionpolicy bypass -file \\area52.afnoapps.usaf.mil\Tyndall_AFB\Logon_Scripts\Stats.ps1")

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

Dim oShell
Dim strCmd
Dim fso
Dim Shortcut
 SET fso = Wscript.CreateObject("Scripting.FileSystemObject")

 Dim FileName
 FileName = "Tyndall AFB Communications Self Help"

 If Not fso.FileExists(FileName + ".lnk") Then
	Dim IconPath 
	IconPath = "C:\Windows\System32\SHELL32.dll, 93"
	Dim TargetPath 
	TargetPath= "https://tyndall.eim.acc.af.mil/SHS/default.aspx"

	Set shortcut = CreateObject("WScript.Shell").CreateShortcut(CreateObject("WScript.Shell").SpecialFolders("Desktop") & + "\" + FileName + ".lnk")

	'Set properties
	shortcut.Description = FileName
	shortcut.TargetPath = TargetPath
	shortcut.IconLocation = IconPath
	shortcut.Save
	
	'Clear Obj ref
	Set shortcut = Nothing   
 End If
Set fso = Nothing

Dim Shortcut2
 SET fso = Wscript.CreateObject("Scripting.FileSystemObject")

 Dim FileName2
 FileName2 = "CHES End User Guide"

 If Not fso.FileExists(FileName2 + ".lnk") Then
	Dim IconPath2 
	IconPath2 = "C:\Windows\System32\SHELL32.dll, 140"
	Dim TargetPath2
	TargetPath2= "https://tyndall.eim.acc.hedc.af.mil/SHS/Connection%20Tester/shs/navigation/CHES_enduserguide.pdf"

	Set shortcut2 = CreateObject("WScript.Shell").CreateShortcut(CreateObject("WScript.Shell").SpecialFolders("Desktop") & + "\" + FileName2 + ".lnk")

	'Set properties
	shortcut2.Description = FileName2
	shortcut2.TargetPath = TargetPath2
	shortcut2.IconLocation = IconPath2
	shortcut2.Save
	
	'Clear Obj ref
	Set shortcut = Nothing   
 End If
Set fso = Nothing

Dim Shortcut3
 SET fso = Wscript.CreateObject("Scripting.FileSystemObject")

 Dim FileName3
 FileName3 = "OCE Tool"

 If Not fso.FileExists(FileName3 + ".lnk") Then
	Dim IconPath3 
	IconPath3 = "C:\Windows\System32\SHELL32.dll, 80"
	Dim TargetPath3
	TargetPath3= "\\xlwu-fs-05pv\Tyndall_PUBLIC\CHES\OCEv3.3-SDC-5.3.1.exe" 

	Set shortcut3 = CreateObject("WScript.Shell").CreateShortcut(CreateObject("WScript.Shell").SpecialFolders("Desktop") & + "\" + FileName3 + ".lnk")

	'Set properties
	shortcut3.Description = FileName3
	shortcut3.TargetPath = TargetPath3
	shortcut3.IconLocation = IconPath3
	shortcut3.Save
	
	'Clear Obj ref
	Set shortcut = Nothing   
 End If
Set fso = Nothing