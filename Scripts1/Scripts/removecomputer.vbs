'--------------------------------------------------------------- 
'                                                                              
'  .:AdamWorkingHard@EvilBytes.com (c) 2006:.      
'                                                                              
'    Name:  AD_DeleteCompObject.vbs    
'    Date:  5/5/2007            
'    Desc:  Uses XLS to deletes computer object from AD 
' Require:  Must have admin rights to delete in AD 
'                                                                              
'--------------------------------------------------------------- 

Option Explicit 

Dim strComputer 
Dim strXLS 
Dim strFunc 
Dim strMsg 
Dim title 

Const adStateOpen = 1 

'----- Call Functions ----- 
Call strInput(strXLS) 'call func 
Call objExcel(strComputer) 'call func 
title = "Systems - Delete Computers" 
strFunc = strXLS 
'-------------------------- 

Function strInput(strXLS) 
'inputbox, grab PATH/FileName.XLS 
strXLS = InputBox(vbLf & vbLf & _ 
"Please enter C:\PATH\FileName.XLS" & vbLf & _ 
"Example:  C:\Folder\Computers.xls" & vbLf & vbLf & vbLf & _ 
"Deletion starts on row 3, column A" & vbLf & _ 
"(A3 downwards)", title) 
'Error trap 
If (IsEmpty(strXLS)) Then 
   strMsg = "Please enter a path" 
   MsgBox strMsg, 48, title 
   Err.Clear 
   wscript.quit 
End If  
End Function 


Function objExcel(strComputer) 
Dim intRow 
Dim objSpread 
'Ensure intRow does not reset to 3 
If Not(IsEmpty(intRow)) Then 
   intRow = intRow + 1 
  Else 
   intRow = 3 'Starting row 
End If 

Set objExcel = CreateObject("Excel.Application") 
Set objSpread = objExcel.Workbooks.Open(strXLS) 
'intRow, x must correspond to the column in strSheet 
Do Until objExcel.Cells(intRow, 1).Value = "" 
   strComputer = Trim(objExcel.Cells(intRow, 1).Value) 'intRow is ROW, 1 = Column 
   Call CompDelete(strComputer) 
   intRow = intRow + 1 'increment row 
Loop 
objExcel.Quit 
End Function 


Function CompDelete(strComputer) 
Dim strDefaultDNC 
Dim strADSQuery 
Dim strAdsPath 
Dim objComp 
Dim objQueryResultSet 
Dim objADOConn 
Dim objADOCommand 
Dim intCount 
'Get the Default Domain Naming Context 
strDefaultDNC = GetObject("LDAP://RootDSE").Get("DefaultNamingContext") 
   If (IsEmpty(strDefaultDNC)) Then 
      Wscript.Echo("Error: Did not get the Default Naming Context") 
      Call Cleanup(1) 
   End If 

'Set up the ADO connection 
Set objADOConn = CreateObject("ADODB.Connection") 
objADOConn.Provider = "ADsDSOObject" 
objADOConn.Open "Active Directory Provider" 'Connect to AD 
   'Verify connection state 
   If objADOConn.State <> adStateOpen Then 
      Wscript.Echo("Authentication Failed.") 
      Call Cleanup(2) 
   End If 

Set objADOCommand = CreateObject("ADODB.Command") 
Set objADOCommand.ActiveConnection = objADOConn 

'Format search of CN(CommonName) using SQL syntax 
strADSQuery = "SELECT * FROM 'LDAP:// " & _ 
strDefaultDNC & "' WHERE CN = '" & strComputer & "'"    
objADOCommand.CommandText = strADSQuery 
Set objQueryResultSet = objADOCommand.Execute 'Execute the search 

'----- Query Check -----  
intCount = 0 
While Not objQueryResultSet.EOF 
   strAdsPath = objQueryResultSet.Fields("AdsPath") 
   intCount = intCount + 1 
   objQueryResultSet.MoveNext 
Wend 
objADOConn.Close 

'Check intCount value. If query 1 result successful 
'If query 0 then failed 
If intCount = 1  Then 
   Set objComp = GetObject(strAdsPath) 
   objComp.DeleteObject (0) 
End If 
End Function 


'---- Sub CleanUp ---- 
Sub Cleanup(intExitCode) 
Dim objADOConn 
  Set objADOConn = Nothing 
  Wscript.Quit(intExitCode) 
End Sub 
'--------------------- 


'----- MSG center ----- 
strMsg = "               AD Comp Audit" & vbLf & _ 
"             Computers deleted" 
MsgBox strMsg, 0, title 
'----------------------- 
