On Error Resume Next
Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
 
Set objUser = GetObject _
    ("LDAP://cn=TINSLEY.JON.K.1041750177.N,ou=area52")
 
intPrimaryGroupID = objUser.Get("primaryGroupID")
arrMemberOf = objUser.GetEx("memberOf")
 
If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
    WScript.Echo "The memberOf attribute is not set."
Else
    WScript.Echo "Member of: "
    For each Group in arrMemberOf
        WScript.Echo Group
    Next
End If
 
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection
objCommand.CommandText = _
  ";(objectCategory=Group);" & _
  "distinguishedName,primaryGroupToken;subtree"  
Set objRecordSet = objCommand.Execute
  
While Not objRecordset.EOF
    If objRecordset.Fields("primaryGroupToken") = intPrimaryGroupID Then
        WScript.Echo "Primary group:"
        WScript.Echo objRecordset.Fields("distinguishedName") & _
            " (primaryGroupID: " & intPrimaryGroupID & ")"
    End If
    objRecordset.MoveNext
Wend
 
objConnection.Close