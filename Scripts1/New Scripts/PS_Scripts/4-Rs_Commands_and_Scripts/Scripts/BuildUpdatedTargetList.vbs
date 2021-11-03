'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: BuildUpdatedTargetList.vbs
'
' AUTHOR: Aaron Mueller
' DATE  : 3/26/2014
'
' COMMENT: Rework of Bill's script
'	Used to generated target file of missing machines based on an original target list compare.
'
'	Version 1.0		Date: 26 Mar 2014
'	Modifications:
'		Initial release of re-written code.
'	Version 1.0.3.0		Date: 25 August 2014
'		Modified version to comply with SCOPE EDGE 4 digit versioning information from Mike West
'	Version 1.0.3.1		Date: 28 November 2014
'		Added function GetRemoteOSVersion to LibraryFunctions Class.
'		Added function LogIt to LibraryFunctions Class.
'	Version 1.0.3.2		Date: 05 December 2014
'		Modified code in function GetActiveIPAddress to use "Route Print" to get the active interface, then check that against
'			all Network Adapter Configurations to find the associated IP Address for the local machine.
'	Version 1.0.3.3		Date: 12 December 2014
'		Added function LogItRecordset to LibraryFunctions Class.
'	Version 1.1.3.4		Date: 20 February 2015
'		Modified function ProcessingLocal (LibraryFunctions Class) to return the local machine name if there was no data passed at startup.
'
'==========================================================================
Option Explicit

'#region <Global Constants>

Const SCRIPT_VERSION_MESSAGE		= "Executing BuildUpdatedTargetList Script Version 1.0.3.4"
Const SCRIPT_VERSION				= "1.0.3.4"
Const DEFAULT_DATE = #1/1/1970#
Const DNS_TIMESTAMP_DATE_BASE = #1/1/1601#

'#endregion

'#region <ADO Constants>

'
' Data type conversion from Access to DAO and ADOX
'
'	Access user interface data type						DAO data type				ADOX data type 
'	Yes/No												dbBoolean					adBoolean 
'	Number or AutoNumber (FieldSize = LongInteger)		dbLong						adInteger 
'	Number (FieldSize = Integer)						dbInteger					adSmallInt 
'	Number (FieldSize = Byte)							dbByte						adUnsignedTinyInt 
'	Number (FieldSize = Decimal)						dbDecimal					adDecimal/adNumeric
'	Number (FieldSize = Double)							dbDouble					adDouble 
'	Number (FieldSize = Single)							dbSingle					adSingle 
'	Number or AutoNumber (FieldSize = Replication ID)	dbGUID						adGUID 
'	Currency											dbCurrency					adCurrency 
'	Binary																			adBinary/adVarBinary
'	OLE Object											dbLongBinary				adLongVarBinary 
'	Text												dbText						adWChar/ADVarWChar 
'	Memo												dbMemo						adLongVarWChar 
'	Date/Time											dbDate						adDate 
'	Hyperlink 											dbMemo, plus DAO			adLongVarWChar, plus ADOX provider-specific
'														Attributes Property set		Column Property set to 
'														to dbHyperlinkField			Jet OLEDB:Hyperlink
'

'
' ADO CursorType Values
'
Const adOpenUnspecified = -1
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3
'
' ADO LockTypeEnum Values
'
Const adLockUnspecified = -1
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4
'
'  ADO DataTypeEnum Values
'
Const adBigInt				= 20	' Indicates an eight-byte signed integer (DBTYPE_I8). 
Const adBinary				= 128	' Indicates a binary value (DBTYPE_BYTES). 
Const adBoolean				= 11	' Indicates a boolean value (DBTYPE_BOOL). 
Const adBSTR				= 8		' Indicates a null-terminated character string (Unicode) (DBTYPE_BSTR). 
Const adChapter				= 136	' Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE_HCHAPTER). 
Const adChar				= 129	' Indicates a string value (DBTYPE_STR). 
Const adCurrency			= 6		' Indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with four digits To
									' the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000. 
Const adDate				= 7		' Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the
									' number of days since December 30, 1899, and the fractional part of which is the fraction of a day. 
Const adDBDate				= 133	' Indicates a date value (yyyymmdd) (DBTYPE_DBDATE). 
Const adDBTime				= 134	' Indicates a time value (hhmmss) (DBTYPE_DBTIME). 
Const adDBTimeStamp			= 135	' Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP). 
Const adDecimal				= 14	' Indicates an exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL). 
Const adDouble				= 5		' Indicates a double-precision floating-point value (DBTYPE_R8). 
Const adEmpty				= 0		' Specifies no value (DBTYPE_EMPTY). 
Const adError				= 10	' Indicates a 32-bit error code (DBTYPE_ERROR). 
Const adFileTime			= 64	' Indicates a 64-bit value representing the number of 100-nanosecond intervals since
									' January 1, 1601 (DBTYPE_FILETIME). 
Const adGUID				= 72	' Indicates a globally unique identifier (GUID) (DBTYPE_GUID). 
Const adIDispatch			= 9		' Indicates a pointer to an IDispatch interface on a COM object (DBTYPE_IDISPATCH). 
									' Note: This data type is currently not supported by ADO. Usage may cause unpredictable results.
Const adInteger				= 3		' Indicates a four-byte signed integer (DBTYPE_I4). 
Const adIUnknown			= 13	' Indicates a pointer to an IUnknown interface on a COM object (DBTYPE_IUNKNOWN). 
									' Note: This data type is currently not supported by ADO. Usage may cause unpredictable results.
Const adLongVarBinary		= 205	' Indicates a long binary value. 
Const adLongVarChar			= 201	' Indicates a long string value. 
Const adLongVarWChar		= 203	' Indicates a long null-terminated Unicode string value. 
Const adNumeric				= 131	' Indicates an exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC). 
Const adPropVariant			= 138	' Indicates an Automation PROPVARIANT (DBTYPE_PROP_VARIANT). 
Const adSingle				= 4		' Indicates a single-precision floating-point value (DBTYPE_R4). 
Const adSmallInt			= 2		' Indicates a two-byte signed integer (DBTYPE_I2). 
Const adTinyInt				= 16	' Indicates a one-byte signed integer (DBTYPE_I1). 
Const adUnsignedBigInt		= 21	' Indicates an eight-byte unsigned integer (DBTYPE_UI8). 
Const adUnsignedInt			= 19	' Indicates a four-byte unsigned integer (DBTYPE_UI4). 
Const adUnsignedSmallInt	= 18	' Indicates a two-byte unsigned integer (DBTYPE_UI2). 
Const adUnsignedTinyInt		= 17	' Indicates a one-byte unsigned integer (DBTYPE_UI1). 
Const adUserDefined			= 132	' Indicates a user-defined variable (DBTYPE_UDT). 
Const adVarBinary			= 204	' Indicates a binary value. 
Const adVarChar				= 200	' Indicates a string value. 
Const adVariant				= 12	' Indicates an Automation Variant (DBTYPE_VARIANT). 
									' Note: This data type is currently not supported by ADO. Usage may cause unpredictable results.
Const adVarNumeric			= 139	' Indicates a numeric value. 
Const adVarWChar			= 202	' Indicates a null-terminated Unicode character string. 
Const adWChar				= 130	' Indicates a null-terminated Unicode character string (DBTYPE_WSTR). 
'
' CursorOptionEnum Values
'
Const adHoldRecords = &H00000100
Const adMovePrevious = &H00000200
Const adAddNew = &H01000400
Const adDelete = &H01000800
Const adUpdate = &H01008000
Const adBookmark = &H00002000
Const adApproxPosition = &H00004000
Const adUpdateBatch = &H00010000
Const adResync = &H00020000
Const adNotify = &H00040000
'
' ExecuteOptionEnum Values
'
Const adRunAsync = &H00000010
'
' ObjectStateEnum Values
'
Const adStateClosed = &H00000000
Const adStateOpen = &H00000001
Const adStateConnecting = &H00000002
Const adStateExecuting = &H00000004
Const adStateFetching = &H00000008
'
' CursorLocationEnum Values
'
Const adUseServer = 2
Const adUseClient = 3
'
' ADO FieldAttributeEnum Values
'
Const adFldMayDefer = &H00000002
Const adFldUpdatable = &H00000004
Const adFldUnknownUpdatable = &H00000008
Const adFldFixed = &H00000010
Const adFldIsNullable = &H00000020
Const adFldMayBeNull = &H00000040
Const adFldLong = &H00000080
Const adFldRowID = &H00000100
Const adFldRowVersion = &H00000200
Const adFldCacheDeferred = &H00001000
'
' ADO EditModeEnum Values
'
Const adEditNone = &H0000
Const adEditInProgress = &H0001
Const adEditAdd = &H0002
Const adEditDelete = &H0004
'
' ADO RecordStatusEnum Values
'
Const adRecOK = &H0000000
Const adRecNew = &H0000001
Const adRecModified = &H0000002
Const adRecDeleted = &H0000004
Const adRecUnmodified = &H0000008
Const adRecInvalid = &H0000010
Const adRecMultipleChanges = &H0000040
Const adRecPendingChanges = &H0000080
Const adRecCanceled = &H0000100
Const adRecCantRelease = &H0000400
Const adRecConcurrencyViolation = &H0000800
Const adRecIntegrityViolation = &H0001000
Const adRecMaxChangesExceeded = &H0002000
Const adRecObjectOpen = &H0004000
Const adRecOutOfMemory = &H0008000
Const adRecPermissionDenied = &H0010000
Const adRecSchemaViolation = &H0020000
Const adRecDBDeleted = &H0040000
'
' ADO GetRowsOptionEnum Values
'
Const adGetRowsRest = -1
'
' ADO PositionEnum Values
'
Const adPosUnknown = -1
Const adPosBOF = -2
Const adPosEOF = -3
'
' ADO BookmarkEnum Values
'
Const adBookmarkCurrent = 0 
Const adBookmarkFirst = 1 
Const adBookmarkLast = 2 
'
' ADO MarshalOptionsEnum Values
'
Const adMarshalAll = 0
Const adMarshalModifiedOnly = 1
'
' ADO AffectEnum Values
'
Const adAffectCurrent = 1
Const adAffectGroup = 2
Const adAffectAll = 3
'
' ADO FilterGroupEnum Values
'
Const adFilterNone = 0
Const adFilterPendingRecords = 1
Const adFilterAffectedRecords = 2
Const adFilterFetchedRecords = 3
Const adFilterPredicate = 4
'
' ADO SearchDirection Values
'
Const adSearchForward = 1
Const adSearchBackward = -1
'
' ADO ConnectPromptEnum Values
'
Const adPromptAlways = 1
Const adPromptComplete = 2
Const adPromptCompleteRequired = 3
Const adPromptNever = 4
'
' ADO ConnectModeEnum Values
'
Const adModeUnknown = 0
Const adModeRead = 1
Const adModeWrite = 2
Const adModeReadWrite = 3
Const adModeShareDenyRead = 4
Const adModeShareDenyWrite = 8
Const adModeShareExclusive = &Hc
Const adModeShareDenyNone = &H10
'
' ADO IsolationLevelEnum Values
'
Const adXactUnspecified = &Hffffffff
Const adXactChaos = &H00000010
Const adXactReadUncommitted = &H00000100
Const adXactBrowse = &H00000100
Const adXactCursorStability = &H00001000
Const adXactReadCommitted = &H00001000
Const adXactRepeatableRead = &H00010000
Const adXactSerializable = &H00100000
Const adXactIsolated = &H00100000
'
' ADO XactAttributeEnum Values
'
Const adXactCommitRetaining = &H00020000
Const adXactAbortRetaining = &H00040000
'
' ADO PropertyAttributesEnum Values
'
Const adPropNotSupported = &H0000
Const adPropRequired = &H0001
Const adPropOptional = &H0002
Const adPropRead = &H0200
Const adPropWrite = &H0400
'
' ADO ErrorValueEnum Values
'
Const adErrInvalidArgument = &Hbb9
Const adErrNoCurrentRecord = &Hbcd
Const adErrIllegalOperation = &Hc93
Const adErrInTransaction = &Hcae
Const adErrFeatureNotAvailable = &Hcb3
Const adErrItemNotFound = &Hcc1
Const adErrObjectInCollection = &Hd27
Const adErrObjectNotSet = &Hd5c
Const adErrDataConversion = &Hd5d
Const adErrObjectClosed = &He78
Const adErrObjectOpen = &He79
Const adErrProviderNotFound = &He7a
Const adErrBoundToCommand = &He7b
Const adErrInvalidParamInfo = &He7c
Const adErrInvalidConnection = &He7d
Const adErrStillExecuting = &He7f
Const adErrStillConnecting = &He81
'
' ADO ParameterAttributesEnum Values
'
Const adParamSigned = &H0010
Const adParamNullable = &H0040
Const adParamLong = &H0080
'
' ADO ParameterDirectionEnum Values
'
Const adParamUnknown = &H0000
Const adParamInput = &H0001
Const adParamOutput = &H0002
Const adParamInputOutput = &H0003
Const adParamReturnValue = &H0004
'
' ADO SchemaEnum Values
'
Const adSchemaProviderSpecific = -1
Const adSchemaAsserts = 0
Const adSchemaCatalogs = 1
Const adSchemaCharacterSets = 2
Const adSchemaCollations = 3
Const adSchemaColumns = 4
Const adSchemaCheckConstraints = 5
Const adSchemaConstraintColumnUsage = 6
Const adSchemaConstraintTableUsage = 7
Const adSchemaKeyColumnUsage = 8
Const adSchemaReferentialContraints = 9
Const adSchemaTableConstraints = 10
Const adSchemaColumnsDomainUsage = 11
Const adSchemaIndexes = 12
Const adSchemaColumnPrivileges = 13
Const adSchemaTablePrivileges = 14
Const adSchemaUsagePrivileges = 15
Const adSchemaProcedures = 16
Const adSchemaSchemata = 17
Const adSchemaSQLLanguages = 18
Const adSchemaStatistics = 19
Const adSchemaTables = 20
Const adSchemaTranslations = 21
Const adSchemaProviderTypes = 22
Const adSchemaViews = 23
Const adSchemaViewColumnUsage = 24
Const adSchemaViewTableUsage = 25
Const adSchemaProcedureParameters = 26
Const adSchemaForeignKeys = 27
Const adSchemaPrimaryKeys = 28
Const adSchemaProcedureColumns = 29
'
' Key type constants
'
Const adKeyPrimary			= 1		' Default. The key is a primary key. 
Const adKeyForeign			= 2		' The key is a foreign key. 
Const adKeyUnique			= 3		' The key is unique. 
'
' Column attributes constants
'
Const adColFixed			= 1		' The column is a fixed length.  
Const adColNullable			= 2		' The column may contain null values. 
'
' ADO CommandTypeEnum Values (Options)
'
Const adCmdText				= 1		' Source holds command text (e.g. a SQL string )
Const adCmdTable			= 2		' Source is the name of a table created by a SQL command (this table is from a database)
Const adCmdStoredProc		= 4		' Source is a stored procedure (e.g. Access stored query) (e.g. objRS.Open "qry_recet_hires",,,,adCmdStoredProc)
Const adCmdUnknown			= 8		' Unknown Source parameter
Const adCmdFile				= 256	' Source is a saved (="persisted") recordset
Const adCmdTableDirect		= 512	' Source is the name of a table (not a database table) ???
'
' Primary Key rules constants
'
Const adRINone				= 0		' Default.  No action is taken. 
Const adRICascade			= 1		' Cascade changes. 
Const adRISetNull			= 2		' Foreign key value is set to NULL.
Const adRISetDefault		= 3		' Foreign key value is set to the default.
'
' AllowNullsEnum Values
'
Const adIndexNullsAllow = 0
Const adIndexNullsDisallow = 1
Const adIndexNullsIgnore = 2
Const adIndexNullsIgnoreAny = 4

'#endregion

'#region <FileIO Constants>

'
' FSO (File I/O) constants
'
Const FOR_READ						= 1
Const FOR_WRITE						= 2
Const FOR_APPEND					= 8
Const CREATE_IF_NON_EXISTENT		= True
Const DONT_CREATE_IF_NON_EXISTENT	= False
Const OVERWRITE_IF_EXISTENT			= True

'#endregion

Dim g_objFSO, g_objFunctions, g_strParentFolder, g_rsXMLFileList, g_rsProcessingList, g_strFolder, g_strFile, g_strArgument
Dim g_objFolder, g_colFiles, g_objFile, g_strFileExtension, g_strMachine, g_strFileName, g_strNewFile, g_objNewFile
Dim g_intNumMachinesToGo, g_intNumMachinesCollected, g_strBuildUpdatedTargetListAudit, g_objTextFile


'
' Classes
'
Class LibraryFunctions
	'
	' Requirements:		objBrowse Class
	'
 	Private m_adVarChar, m_adStateOpen, m_DEFAULT_DATE, m_FOR_WRITE, m_OVERWRITE_IF_EXISTENT, m_OS_VERSION_XP
 	Private m_INTEGER8_DATE_NOADJUST, m_INTEGER8_DATE_ADJUST_LOCAL, m_INTEGER8_DATE_ADJUST_GMT, m_ADS_SCOPE_SUBTREE
	Private m_ADS_CHASE_REFERRALS_ALWAYS, m_vbGUID, m_vbSID, m_vbScheduleOrRelay, m_vbDateDNSTimestamp, m_vbDateMicrosoftDateTime
	Private m_vbDateWMIDateTime, m_vbDateUTCDateTime, m_DNS_TIMESTAMP_DATE_BASE, m_MSSQL_TIMESTAMP_DATE_BASE, m_HKEY_LOCAL_MACHINE
	Private m_ADS_SERVER_BIND, m_ADS_SECURE_AUTHENTICATION


	Private Sub Class_Initialize() 'Constructor
 		m_adVarChar = 200	' Indicates a string value. 
 		m_adStateOpen = 1
 		m_DEFAULT_DATE = #1/1/1970#
 		m_FOR_WRITE = 2
 		m_OVERWRITE_IF_EXISTENT = True
 		m_OS_VERSION_XP = "5.1"
 		m_INTEGER8_DATE_NOADJUST = 1
 		m_INTEGER8_DATE_ADJUST_LOCAL = 2
 		m_INTEGER8_DATE_ADJUST_GMT = 3
 		m_ADS_SCOPE_SUBTREE = 2
		m_ADS_CHASE_REFERRALS_ALWAYS = &H60 
		m_vbGUID = 100
		m_vbSID = 101
		m_vbScheduleOrRelay = 102
		m_vbDateDNSTimestamp = 103
		m_vbDateMicrosoftDateTime = 104
		m_vbDateWMIDateTime = 105
		m_vbDateUTCDateTime = 106
		m_DNS_TIMESTAMP_DATE_BASE = #1/1/1601#
		m_MSSQL_TIMESTAMP_DATE_BASE	= #1/1/1801#
		m_HKEY_LOCAL_MACHINE = &H80000002
		m_ADS_SECURE_AUTHENTICATION = &H1
		m_ADS_SERVER_BIND = &H200
	End Sub

	Private Sub Class_Terminate 'Destructor
    End Sub

	Private Sub LogThis(ByVal strText, ByRef objLogAndTrace)
		Dim strTextLocal
		If (IsObject(objLogAndTrace)) Then
			strTextLocal = strText
			Call CreatePrintableString(strTextLocal)
			objLogAndTrace.LogThis(strTextLocal)
		End If
	End Sub

	Public Function LogIt(ByVal xmlElement, ByRef objLogAndTrace)
	'*****************************************************************************************************************************************
	'*  Purpose:				Writes data from the xmlElement to the LogAndTrace object
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					LogThis
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim intCount, strAttribute, varValue
	
		For intCount = 0 To xmlElement.Attributes.length - 1
			strAttribute = xmlElement.Attributes.Item(intCount).nodeName
			varValue = xmlElement.Attributes.Item(intCount).nodeValue
			objLogAndTrace.LogThis(strAttribute & ": " & varValue)
		Next
		objLogAndTrace.LogThis("")
	
	End Function

	Public Function LogItRecordset(ByRef rsToLog, ByRef objLogAndTrace)
	'*****************************************************************************************************************************************
	'*  Purpose:				Writes data from rsToLog to the LogAndTrace object
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					LogThis
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim intCount, strColumnName, intCreateDataType, intCreateMaxLength

		If (rsToLog.Fields.Count > 0) Then
			For intCount = 0 To rsToLog.Fields.Count - 1
				strColumnName = rsToLog(intCount).Name
				intCreateDataType = rsToLog(intCount).Type
				intCreateMaxLength = rsToLog(intCount).DefinedSize
				objLogAndTrace.LogThis(strColumnName & ": " & rsToLog(intCount))
'				WScript.Echo strColumnName & ": " & rsToLog(intCount)
			Next
		End If

	End Function

	Public Function LPad(ByVal strValue, ByVal intStringLength, ByVal strPadCharacter)
	'*****************************************************************************************************************************************
	'*  Purpose:				Left-pad an input value to the given number of characters using the given padding character
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim intPadCharacters
		
		intPadCharacters = 0
		If (intStringLength > Len(strValue)) Then
			intPadCharacters = intStringLength - Len(strValue)
		End If
		LPad = String(intPadCharacters, strPadCharacter) & strValue

	End Function

	Public Function ConvertToDNS(ByVal strDistinguishedName)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts a DistinguishedName (Domain) to DomainDNSName.
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************

		ConvertToDNS = Replace(Replace(strDistinguishedName,"DC=",""),",",".")

	End Function
	
	Public Function ConvertToDN(ByVal strDomainDNSName)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts a DomainDNSName to DistinguishedName (Domain).
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************

		ConvertToDN = "DC=" & Replace(strDomainDNSName, ".", ",DC=")

	End Function

	Public Function ProcessingLocal(ByRef strPassedParam)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if the processing is occurring on the local machine
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 if successful, -1 if unsuccessful.
	'*  Called by:				Mainline
	'*  Calls:					None
	'*  Requirements:			ADO Constants
	'*****************************************************************************************************************************************
		Dim strHostName, strIPv4Address, rsGeneric, intRetVal, strTemp, strWork, objShell, strThisComputer
	
		'
		' Get local network information
		'
		strHostName = ""
		strIPv4Address = ""
		
		Set rsGeneric = CreateObject("ADODB.Recordset")
		rsGeneric.Fields.Append "SavedData", m_adVarChar, 255
		rsGeneric.Open
		
		intRetVal = ExecCmdGeneric("ipconfig /all", rsGeneric, "")
		If (intRetVal = 0) Then
			'
			' Got something back
			'
			If (Not rsGeneric.BOF) Then
				rsGeneric.MoveFirst
			End If
			While Not rsGeneric.EOF
				strTemp = rsGeneric("SavedData")
				If (InStr(1, strTemp, "Host Name", vbTextCompare) > 0) Then
					strHostName = Trim(Split(strTemp, ":", 2)(1))
				End If
				If (InStr(1, strTemp, "IPv4 Address", vbTextCompare) > 0) Then
					strWork = Replace(Trim(Split(strTemp, ":", 2)(1)), "(Preferred)", "", 1, 1, vbTextCompare)
					If (strIPv4Address = "") Then
						strIPv4Address = strWork
					Else
						strIPv4Address = strIPv4Address & ";" & strWork
					End If
				End If
				rsGeneric.MoveNext
			Wend
		End If
		'
		' Local processing?
		'
		ProcessingLocal = False
		Set objShell = CreateObject("WScript.Shell")
		strThisComputer = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
		If (strPassedParam = "") Then
			ProcessingLocal = True
			strPassedParam = strThisComputer
		Else
			If ((InStr(1, strPassedParam, strThisComputer, vbTextCompare) > 0) Or _
				(InStr(1, strPassedParam, strHostName, vbTextCompare) > 0) Or _
				(InStr(1, strIPv4Address, strPassedParam, vbTextCompare) > 0)) Then
				ProcessingLocal = True
			End If
		End If
		'
		' Cleanup
		'
		Set rsGeneric = Nothing
		Set objShell = Nothing
	
	End Function

	Public Function CheckRegistryAccess(ByVal objReg, ByVal strHive, ByVal strKey, ByVal strAccessLevel)
	'*****************************************************************************************************************************************
	'*  Purpose:				Checks to see if we have access to the specified registry location
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				GatherMachineData
	'*  Calls:					None
	'*  Requirements:			None
	'*  Example:	CheckRegistryAccess(objReg, HKCU, "Software\KillerApp", KEY_CREATE_SUB_KEY)
	'*					where strAccessLevel is one of the following:
	'*						REG_KEY_QUERY
	'*						REG_KEY_SET
	'*						KEY_CREATE_SUB_KEY
	'*						KEY_ENUMERATE_SUB_KEYS
	'*						KEY_NOTIFY
	'*						KEY_CREATE_LINK
	'*						KEY_DELETE
	'*						READ_CONTROL
	'*						WRITE_DAC
	'*						WRITE_OWNER
	'*
	'*****************************************************************************************************************************************
		Dim blnValue, intErrNumber
		
		'
		' CheckAccess will return true if user has specified permissions
		'
		CheckRegistryAccess = False
		On Error Resume Next
		objReg.CheckAccess strHive, strKey, strAccessLevel, blnValue
		intErrNumber = Err.Number
		On Error GoTo 0
		If (intErrNumber = 0) Then
			CheckRegistryAccess = blnValue
		End If
		
	End Function

	Public Function GenerateRandomPassword(ByVal intLength)
	'*****************************************************************************************************************************************
	'*  Purpose:				Generates a random password of intLength characters.
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Any wishing to create a random password
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim intHighNumber, intLowNumber, intCount, intNumber
		
		intHighNumber = 126
		intLowNumber = 32
		GenerateRandomPassword = ""

		For intCount = 0 To intLength - 1
		    Randomize
		    intNumber = Int((intHighNumber - intLowNumber + 1) * Rnd + intLowNumber)
	    	GenerateRandomPassword = GenerateRandomPassword & Chr(intNumber)
		Next

	End Function

	Public Function RemoveControlCharacters(ByVal strToCheck)
	'*****************************************************************************************************************************************
	'*  Purpose:				Checks the character string and removes any control characters
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Any wishing to check for valid characters
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strWork, intCount, strChar, strNew
		
		strWork = strToCheck
		'
		' Replace an opening single-tick from Office (Chr(145)) with an ascii single-tick (Chr(39))
		' Replace a closing single-tick from Office (Chr(146)) with an ascii single-tick (Chr(39))
		' Replace an opening double-dash from Office (Chr(147)) with an ascii double-dash (Chr(34))
		' Replace a closing double-dash from Office (Chr(148)) with an ascii double-dash (Chr(34))
		' Replace a bullet character from Office (Chr(149)) with an asterisk (Chr(42))
		' Replace a dash from Office (Chr(150)) with an ascii dash (Chr(45))
		'
		strWork = Replace(strWork, Chr(145), Chr(39))
		strWork = Replace(strWork, Chr(146), Chr(39))
		strWork = Replace(strWork, Chr(147), Chr(34))
		strWork = Replace(strWork, Chr(148), Chr(34))
		strWork = Replace(strWork, Chr(149), Chr(42))
		strWork = Replace(strWork, Chr(150), Chr(45))
		'
		' Remove all of the control characters.  The Unicode standard defines 0x00-0x1F (00-31 Decimal), 0x7F (127 Decimal), and 
		' 0x80-0x9F (128-159 Decimal) as control characters.
		'
		For intCount = 0 To Len(strWork) - 1
			strChar = Mid(strWork, intCount + 1, 1)
			If ((Asc(strChar) >= 32 And Asc(strChar) <= 126) Or (Asc(strChar) >= 160)) Then
				strNew = strNew & strChar
			End If
		Next
		RemoveControlCharacters = strNew

	End Function

	Public Function ContainsValidCharacters(ByVal strToCheck)
	'*****************************************************************************************************************************************
	'*  Purpose:				Checks the character string to ensure it only contains valid characters
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Any wishing to check for valid characters
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim intCount, strChar
		
		If ((IsNull(strToCheck)) Or (Len(strToCheck) = 0)) Then
			ContainsValidCharacters = -1
			Exit Function
		End If

		For intCount = 0 To Len(strToCheck) - 1
			strChar = Mid(strToCheck, intCount + 1, 1)
			If (Asc(strChar) < 32 Or (Asc(strChar) > 126 And Asc(strChar) < 160)) Then
				ContainsValidCharacters = -1
				Exit Function
			End If
		Next
		ContainsValidCharacters = 0

	End Function

	Public Function SaveRecordset(ByVal objRS, ByVal strRSFile)
	'*****************************************************************************************************************************************
	'*  Purpose:				Saves the recordset to a file
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All functions that want to save a recordset (DUH!!!)
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim objFSO, strSaveFile
		
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		'
		' Save the recordset (only write if there is data)
		'
		If (objRS.RecordCount > 0) Then
			If (objFSO.FileExists(strRSFile)) Then
				objFSO.DeleteFile(strRSFile)
			End If
			objRS.Save strRSFile
		End If
		'
		' Cleanup
		'
		Set objFSO = Nothing
	
	End Function

	Public Function GetADInfo(ByRef rsAD, ByVal strSQLQuery, ByVal intSearchScope, ByRef intErrNumber, ByRef strErrDescription, _
									ByRef objLogAndTrace, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Executes the specified SQL query/queries against AD
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Main()
	'*  Calls:					None
	'*  Requirements:			ADO Constants
	'*****************************************************************************************************************************************
		Dim objConnection, objCommand

		'
		' Create a Connection object in memory and open the Connection object using the ADSI OLE DB provider.
		'
		Set objConnection = CreateObject("ADODB.Connection")
		objConnection.Open "Provider=ADsDSOObject;"
		'
		' Create an ADO Command object in memory and assign the Command object's ActiveConnection property to the Connection object. 
		'
		Set objCommand = CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConnection
		objCommand.CommandText = strSQLQuery
		objCommand.Properties("Page Size") = 20000
		objCommand.Properties("Timeout") = 20000
'		objCommand.Properties("Searchscope") = m_ADS_SCOPE_SUBTREE
		objCommand.Properties("Searchscope") = intSearchScope
		objCommand.Properties("Cache Results") = True
'		objCommand.Properties("Deref Aliases") = True 
'		objCommand.Properties("Encrypt Password") = True 
		objCommand.Properties("Chase referrals") = m_ADS_CHASE_REFERRALS_ALWAYS

		On Error Resume Next
		Set rsAD = objCommand.Execute
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
		'
		' Log information about query results
		'
		Call LogThis("In GetADInfo", objLogAndTrace)
		Call LogThis("SQLQuery: " & strSQLQuery, objLogAndTrace)
		Call LogThis("Error: " & intErrNumber & vbTab & "Description: " & strErrDescription, objLogAndTrace)
		Call LogThis("State: " & rsAD.State, objLogAndTrace)
		If (rsAD.State = m_adStateOpen) Then
			Call LogThis("RecordCount: " & rsAD.RecordCount, objLogAndTrace)
		End If
		'
		' Check to see if query returned anything
		'
		If (intErrNumber <> 0) Then
			'
			' An Error occurred or Recordset not open (error)
			'
			Call LogThis("GetADInfo processing failed.  Error: " & intErrNumber & "  " & "Description: " & strErrDescription, objLogAndTraceErrors)
			GetADInfo = -1
		Else
			'
			' No error occurred - check state
			'
			If (rsAD.State <> m_adStateOpen) Then
				'
				' Database couldn't be opened - error
				'
				Call LogThis("GetADInfo processing failed - Database couldn't be opened", objLogAndTraceErrors)
				GetADInfo = -1
			Else
				If (rsAD.RecordCount<>0) Then
					GetADInfo = 0
				Else
					Call LogThis("GetADInfo processing failed - No records returned from AD", objLogAndTrace)
					GetADInfo = -1
				End If
			End If
		End If

	End Function

	Public Function DoesColumnExist(ByRef rsColumns, ByVal strColumnName, ByRef intDataType, ByRef intMaxLength)
	'*****************************************************************************************************************************************
	'*  Purpose:				Checks to see if this column exists
	'*  Arguments supplied:		None
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					DeleteAllRecordsetRows
	'*  Requirements:			ADO Constants
	'*****************************************************************************************************************************************
		Dim intCount, strName, intErrNumber
	
		intDataType = 0
		intMaxLength = 0
	' 	For intCount = 0 To rsColumns.Fields.Count - 1
	' 		If (InStr(1, rsColumns(intCount).Name, strColumnName) > 0) Then
	' 			intDataType = rsColumns(intCount).Type
	' 			intMaxLength = rsColumns(intCount).DefinedSize
	' 			DoesColumnExist = True
	' 			Exit Function
	' 		End If
	' 	Next
	' 	DoesColumnExist = False
	
		On Error Resume Next
		strName = rsColumns(strColumnName).Name
		intDataType = rsColumns(strColumnName).Type
		intMaxLength = rsColumns(strColumnName).DefinedSize
		intErrNumber = Err.Number
		On Error GoTo 0
		If (intErrNumber = 0) Then
			DoesColumnExist = True
		Else
			DoesColumnExist = False
		End If
	
	End Function

	Public Function GetAvailableDriveLetter(ByRef objWMIService, ByVal intFlag, ByRef strDriveLetter)
	'*****************************************************************************************************************************************
	'*  Purpose:				Returns the next available drive letter
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim arrDriveLetters, intCount, strSQLQuery, intErrNumber, strErrDescription, colWMI
		
		arrDriveLetters = Array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z")

		For intCount = 0 To UBound(arrDriveLetters)
			strSQLQuery = "SELECT Name FROM Win32_LogicalDisk WHERE DeviceID = '" & arrDriveLetters(intCount) & ":'"
			Call ExecWMI(objWMIService, intErrNumber, strErrDescription, colWMI, strSQLQuery, intFlag, Null)
'			WScript.Echo intErrNumber & vbTab & strErrDescription & vbTab & TypeName(colWMI)
			If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
				If (colWMI.Count=0) Then
					strDriveLetter = arrDriveLetters(intCount) & ":"
					GetAvailableDriveLetter = True
					Exit Function
				End If
			End If
		Next
		'
		' This shouldn't happen unless all 26 drive letters are being used or there is a problem with the WMI results
		'
		GetAvailableDriveLetter = False

	End Function

	Public Function GetUserNameFromSID(ByVal objWMIService, ByVal strSID, ByRef strUser, ByRef strDomain)
	'*****************************************************************************************************************************************
	'*  Purpose:				Gets the user name (and domain) associated with the specified SID
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim objAccount, intErrNumber, strErrDescription
		
		On Error Resume Next
		Set objAccount = objWMIService.Get("Win32_SID.SID='" & strSID & "'")
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
			
		If (intErrNumber = 0) Then
			strUser = objAccount.AccountName
			strDomain = objAccount.ReferencedDomainName
		Else
			strUser = ""
			strDomain = ""
		End If
	
	End Function

	Public Function ExecWMI(ByRef objWMIService, ByRef intErrNumber, ByRef strErrDescription, ByRef colWMI, ByVal strSQLQuery, ByVal intFlag, _
								ByVal objNamedValueSet)
	'*****************************************************************************************************************************************
	'*  Purpose:				Calls WMI ExecQuery and populates all required values
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
	
		On Error Resume Next
		If (objNamedValueSet = Null) Then
			Set colWMI = objWMIService.ExecQuery(strSQLQuery,, intFlag)
		Else
			Set colWMI = objWMIService.ExecQuery(strSQLQuery,, intFlag, objNamedValueSet)
		End If
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
	
	End Function

	Public Function GetMySiteGUID(ByRef rsSubnets, ByVal bigintIPAddress)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines which site the bigintIPAddress falls within
	'*  Arguments supplied:		Look up
	'*  Return Value:			SiteName
	'*  Called by:				Any
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim bigintIPAddressStart, bigintIPAddressEnd, strSiteGUID, Temp

		If (rsSubnets.RecordCount = 0) Then
			GetMySiteGUID = "Unknown"
			Exit Function
		End If
		'
		' Start with the most restrictive subnets then move to more general subnets
		'
		rsSubnets.Sort = "sneMaskBits DESC"
		If (Not rsSubnets.BOF) Then
			rsSubnets.MoveFirst
		End If
		While Not rsSubnets.EOF
			bigintIPAddressStart = rsSubnets("sneIPAddressRangeStartBigInt")
			bigintIPAddressEnd = rsSubnets("sneIPAddressRangeEndBigInt")
			strSiteGUID = rsSubnets("sneSiteGUID")
			If ((CDbl(bigintIPAddress) >= CDbl(bigintIPAddressStart)) And (CDbl(bigintIPAddress) <= CDbl(bigintIPAddressEnd))) Then
				GetMySiteGUID = strSiteGUID
				rsSubnets.Sort = ""
				Exit Function
			End If
			rsSubnets.MoveNext
		Wend
		'
		' Site not found
		'
		GetMySiteGUID = "Unknown"
		rsSubnets.Sort = ""

	End Function

	Public Function GetMySiteObjectGUID(ByRef rsSubnets, ByVal bigintIPAddress)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines which site the bigintIPAddress falls within
	'*  Arguments supplied:		Look up
	'*  Return Value:			SiteName
	'*  Called by:				Any
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim bigintIPAddressStart, bigintIPAddressEnd, strSiteObjectGUID, Temp

		If (rsSubnets.RecordCount = 0) Then
			GetMySiteObjectGUID = "Unknown"
			Exit Function
		End If
		If (Not rsSubnets.BOF) Then
			rsSubnets.MoveFirst
		End If
		While Not rsSubnets.EOF
			bigintIPAddressStart = rsSubnets("sneIPAddressRangeStartBigInt")
			bigintIPAddressEnd = rsSubnets("sneIPAddressRangeEndBigInt")
			strSiteObjectGUID = rsSubnets("sneSiteObjectGUID")
			If ((CDbl(bigintIPAddress) >= CDbl(bigintIPAddressStart)) And (CDbl(bigintIPAddress) <= CDbl(bigintIPAddressEnd))) Then
				GetMySiteObjectGUID = strSiteObjectGUID
				Exit Function
			End If
			rsSubnets.MoveNext
		Wend
		'
		' Site not found
		'
		GetMySiteObjectGUID = "Unknown"

	End Function

	Public Function AmIAnExchangeServer(ByVal strDCHostDNSName, ByVal strComputerName)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if this machine is an Exchange server (according to AD)
	'*  Arguments supplied:		Look up
	'*  Return Value:			True if Exchange Server, False if not
	'*  Called by:				None
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim rsExchAD, objRootDSE, intErrNumber, strErrDescription, strConnection, strSQLQuery
	
		Set rsExchAD = CreateObject("ADODB.Recordset")
		On Error Resume Next
		Set objRootDSE = GetObject("LDAP://" & strDCHostDNSName & "/RootDSE")
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
		If (intErrNumber <> 0) Then
			AmIAnExchangeServer = False
			Exit Function
		End If
		strConnection = "Provider=ADsDSOObject"
		strSQLQuery = "SELECT ADsPath,distinguishedName,cn FROM 'LDAP://" & objRootDSE.Get("configurationNamingContext") & _
						"' WHERE objectCategory='msExchExchangeServer' and cn='" & Replace(strComputerName, "'", "''") & "'"
		On Error Resume Next
		rsExchAD.Open strSQLQuery, strConnection
		intErrNumber = Err.Number
		On Error GoTo 0
		If (intErrNumber = 0) Then
			If (rsExchAD.RecordCount > 0) Then
				AmIAnExchangeServer = True
			Else
				AmIAnExchangeServer = False
			End If
		Else
			AmIAnExchangeServer = False
		End If
	
	End Function

	Public Function FormatMACAddress(ByVal strMACAddress)
	'*****************************************************************************************************************************************
	'*  Purpose:				Removes ":" from MACAddress and makes uppercase.
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Any
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strFormattedMACAddress

		strFormattedMACAddress = UCase(strMACAddress)
		If (InStr(strMACAddress, ":") > 0) Then
			strFormattedMACAddress = Replace(strFormattedMACAddress, ":", "")
		End If
		If (InStr(strMACAddress, "-") > 0) Then
			strFormattedMACAddress = Replace(strFormattedMACAddress, "-", "")
		End If
		FormatMACAddress = strFormattedMACAddress

	End Function

	Public Function MassageTimestamp(ByVal dtToProcess)
	'*****************************************************************************************************************************************
	'*  Purpose:				Replaces "." with "/" in the date string
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Any
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim dtWork, intDay, intMonth, intYear, intHour, intMinute, intSecond, strMonth, strDay, strYear, strAMPM, strHour, strMinute, strSecond

		dtWork = dtToProcess
		If ((IsNull(dtToProcess)) Or (Len(dtToProcess) = 0)) Then
			MassageTimestamp = DEFAULT_DATE
			Exit Function
		End If
		'
		' See if time format contains periods ("03.13.2013 7:35:07 PM") - different date format if running on the client machine
		'
		If (InStr(dtWork, ".") > 0) Then
			dtWork = Replace(dtWork, ".", "/")
		End If
		'
		' See if time format contains dashes ("03-13-2013 7:35:07 PM") - different date format if running on the client machine
		'
		If (InStr(dtWork, "-") > 0) Then
			dtWork = Replace(dtWork, "-", "/")
		End If
		dtWork = Replace(dtWork, "1/1/1601", "1/1/1801")

		intDay = Day(dtWork)
		intMonth = Month(dtWork)
		intYear = Year(dtWork)
		intHour = Hour(dtWork)
		intMinute = Minute(dtWork)
		intSecond = Second(dtWork)
		'
		' Format Month
		'
		If (intMonth < 10) Then
			strMonth = "0" & CStr(intMonth)
		Else
			strMonth = CStr(intMonth)
		End If
		'
		' Format Day
		'
		If (intDay < 10) Then
			strDay = "0" & CStr(intDay)
		Else
			strDay = CStr(intDay)
		End If
		'
		' Format Year
		'
		strYear = CStr(intYear)
		'
		' Determine if the Time (in 24-hour format) is AM or PM
		'
		If ((intHour >= 0) And (intHour <= 11)) Then
			strAMPM = "AM"
		Else
			strAMPM = "PM"
		End If
		'
		' Convert hour to standard time vs. military time
		'
		If (intHour > 12) Then
			intHour = intHour - 12
		End If
		'
		' Adjust any Time values that are 0 to contain 2 digits
		'
		If (intHour < 10) Then
			strHour = "0" & CStr(intHour)
		Else
			strHour = CStr(intHour)
		End If
		If (intMinute < 10) Then
			strMinute = "0" & CStr(intMinute)
		Else
			strMinute = CStr(intMinute)
		End If
		If (intSecond < 10) Then
			strSecond = "0" & CStr(intSecond)
		Else
			strSecond = CStr(intSecond)
		End If
		MassageTimestamp = strMonth & "/" & strDay & "/" & strYear & " " & strHour & ":" & strMinute & ":" & strSecond & " " & strAMPM

	End Function

	Public Function GetGMTTimestamp()
	'*****************************************************************************************************************************************
	'*  Purpose:				Gets the current time and modifies it to be GMT
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Any
	'*  Calls:					MassageTimestamp
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim objShell, strRegistry, intOffsetInMinutes, dtNow
		
		Set objShell = CreateObject("WScript.Shell")
		strRegistry = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
		intOffsetInMinutes = objShell.RegRead(strRegistry)
		dtNow = DateAdd("n", intOffsetInMinutes, Now())
'		Call MassageTimestamp(dtNow)
		GetGMTTimestamp = MassageTimestamp(dtNow)
		Set objShell = Nothing

	End Function

	Public Function GetADPropertyList(ByVal strADsPath)
	'*****************************************************************************************************************************************
	'*  Sub Name:                 GetADPropertyList()
	'*  Purpose:                  Echos the properties/attributes for a specified ADsPath.
	'*  Arguments supplied:       ADsPath
	'*  Return Value:             none
	'*  Sub is called by:         none
	'*  Sub calls:                none
	'*****************************************************************************************************************************************
		Dim objADs, objADsPropList, objADsPropEntry, intPropCount, intPropIndex
			
	'	strADsPath = "LDAP://vejxamcw2dc102.amc.ds.af.mil"
	
		Set objADs = GetObject(strADsPath)
		objADs.GetInfo
		Set objADsPropList = objADs
		intPropCount = objADsPropList.PropertyCount
	
		'Display some information
		WScript.Echo "Object " & objADs.Name & " at " & objADs.ADsPath
		WScript.Echo "Cache contains " & intPropCount & " items"
	
		For intPropIndex = 0 To intPropCount - 1
			'
			' The Item method accepts a text name or index number and returns a IADsPropertyEntry object
			'
			Set objADsPropEntry = objADsPropList.Item(intPropIndex)
			WScript.Echo "Name/Type/Code: " & objADsPropEntry.Name & " / " & objADsPropEntry.ADsType & " / "
		Next
	
	End Function

	Public Function ParseName(ByVal strToParse, ByVal intParseComponent)
	'*****************************************************************************************************************************************
	'*  Purpose:				Parses a string (i.e. an IPv4 address) and returns the requested 'component'
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				IPAddress2BigInt
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim arrElements, intElementCount
	
		If (InStr(strToParse, ".") = 0) Then
			ParseName = Null
			Exit Function
		End If
		'
		' ParseName will only work on a max of 4 elements (1-4)
		'
		If ((intParseComponent < 1) Or (intParseComponent > 4)) Then
			ParseName = Null
			Exit Function
		End If
		'
		' Get the elements ready to process
		'
		arrElements = Split(strToParse, ".")
		intElementCount = UBound(arrElements)
		'
		' Make sure we have the element requested
		'
		If ((intParseComponent - 1) > intElementCount) Then
			ParseName = Null
			Exit Function
		End If
		'
		' We have a valid request and (hopefully) a valid string
		'
		ParseName = CLng(arrElements(intElementCount - (intParseComponent - 1)))
	
	End Function
	
	Public Function IPAddress2BigInt(ByVal strIPAddress)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts an IPv4 address to BigInteger
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					ParseName
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim intRetVal, biConvertedIPAddress
		
		If (InStr(strIPAddress, ".") = 0) Then
			IPAddress2BigInt = 0
			Exit Function
		End If
		'
		' Get first octet
		'
		intRetVal = ParseName(strIPAddress, 1)
		If (IsNull(intRetVal)) Then
			IPAddress2BigInt = 0
			Exit Function
		End If
		If ((intRetVal < 0) Or (intRetVal > 255)) Then
			IPAddress2BigInt = 0
			Exit Function
		End If
		biConvertedIPAddress = biConvertedIPAddress + intRetVal
		'
		' Get second octet
		'
		intRetVal = ParseName(strIPAddress, 2)
		If (IsNull(intRetVal)) Then
			IPAddress2BigInt = 0
			Exit Function
		End If
		If ((intRetVal < 0) Or (intRetVal > 255)) Then
			IPAddress2BigInt = 0
			Exit Function
		End If
		biConvertedIPAddress = biConvertedIPAddress + (intRetVal * 256)
		'
		' Get third octet
		'
		intRetVal = ParseName(strIPAddress, 3)
		If (IsNull(intRetVal)) Then
			IPAddress2BigInt = 0
			Exit Function
		End If
		If ((intRetVal < 0) Or (intRetVal > 255)) Then
			IPAddress2BigInt = 0
			Exit Function
		End If
		biConvertedIPAddress = biConvertedIPAddress + (intRetVal * 65536)
		'
		' Get fourth (final) octet
		'
		intRetVal = ParseName(strIPAddress, 4)
		If (IsNull(intRetVal)) Then
			IPAddress2BigInt = 0
			Exit Function
		End If
		If ((intRetVal < 0) Or (intRetVal > 255)) Then
			IPAddress2BigInt = 0
			Exit Function
		End If
		biConvertedIPAddress = biConvertedIPAddress + (intRetVal * 16777216)
		IPAddress2BigInt = biConvertedIPAddress
	
	End Function

	Public Function BigInt2IPAddress(ByVal bigIntIPAddress)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts a BigInteger to IPv4 address
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					ParseName
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim bigIntWork, intFirstOctet, intSecondOctet, intThirdOctet, intFourthOctet

		bigIntWork = bigIntIPAddress
		'
		' Get first octet
		'
		intFirstOctet = Int(bigIntWork / 16777216)
		If ((intFirstOctet < 0) Or (intFirstOctet > 255)) Then
			BigInt2IPAddress = ""
			Exit Function
		End If
'		WScript.Echo "intFirstOctet: " & intFirstOctet
		bigIntWork = bigIntWork - (intFirstOctet * 16777216)
'		WScript.Echo "bigIntWork: " & bigIntWork
		'
		' Get second octet
		'
		intSecondOctet = Int(bigIntWork / 65536)
		If ((intSecondOctet < 0) Or (intSecondOctet > 255)) Then
			BigInt2IPAddress = ""
			Exit Function
		End If
'		WScript.Echo "intSecondOctet: " & intSecondOctet
		bigIntWork = bigIntWork - (intSecondOctet * 65536)
'		WScript.Echo "bigIntWork: " & bigIntWork
		'
		' Get third octet
		'
		intThirdOctet = Int(bigIntWork / 256)
		If ((intThirdOctet < 0) Or (intThirdOctet > 255)) Then
			BigInt2IPAddress = ""
			Exit Function
		End If
'		WScript.Echo "intThirdOctet: " & intThirdOctet
		bigIntWork = bigIntWork - (intThirdOctet * 256)
		'
		' Get fourth (final) octet
		'
		intFourthOctet = bigIntWork
		If ((intFourthOctet < 0) Or (intFourthOctet > 255)) Then
			BigInt2IPAddress = ""
			Exit Function
		End If
'		WScript.Echo "intFourthOctet: " & intFourthOctet
		BigInt2IPAddress = intFirstOctet & "." & intSecondOctet & "." & intThirdOctet & "." & intFourthOctet

	End Function

	Public Function SetBit(ByRef intBitMap, ByRef intBitToSet)
	'*****************************************************************************************************************************************
	'*  Purpose:				Sets bit 0-31 of a 32 bit integer
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				LoadIntersiteTransportTable
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
	
		If (intBitToSet = 31) Then
			intBitMap = CLng(intBitMap) * -1
		Else
			intBitMap = CLng(intBitMap) + CLng(2^intBitToSet)
		End If
	
	End Function

	Public Function DisplayBitMap(ByVal intBitMap)
	'*****************************************************************************************************************************************
	'*  Purpose:				Displays bits 0-31 of a 32 bit integer
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				None
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strBitMap, intValue, blnHighOrderBitSet, intCount, intBitOffsetValue
		
		strBitMap = ""
		intValue = intBitMap
		blnHighOrderBitSet = False
		If (intValue < 0) Then
			'
			' Negative number
			'
			blnHighOrderBitSet = True
			intValue = intValue * -1
		End If
		For intCount = 0 To 30 Step 1
			intBitOffsetValue = 2^intCount
			If (intValue And intBitOffsetValue) Then
				strBitMap = strBitMap & "1"
				intValue = intValue - intBitOffsetValue
			Else
				strBitMap = strBitMap & "0"
			End If
		Next
		If (blnHighOrderBitSet) Then
			strBitMap = strBitMap & "1"
		Else
			strBitMap = strBitMap & "0"
		End If
		DisplayBitMap = strBitMap
		
	End Function

	Public Function CreateGloballyUniqueID()
	'*****************************************************************************************************************************************
	'*  Purpose:				Creates a GUID
	'*  Arguments supplied:		Look up
	'*  Return Value:			Newly created GUID
	'*  Called by:				All
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim TypeLib, strGUID
		
		Set TypeLib = CreateObject("Scriptlet.TypeLib")
		strGUID = UCase(TypeLib.Guid)
		
		strGUID = Replace(Replace(strGUID, "{", ""), "}", "")
		strGUID = Left(strGUID, 36)
		'
		' Cleanup
		'
		Set TypeLib = Nothing
		CreateGloballyUniqueID = strGUID
	
	End Function

	Public Function TrueOrFalse(ByVal blnValue)
	'*****************************************************************************************************************************************
	'*  Purpose:				Returns the correct value to load in a SQL database for True/False
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
	
		If (blnValue) Then
			TrueOrFalse = 1
		Else
			TrueOrFalse = 0
		End If
	
	End Function

	Public Function Min(ByVal intVal1, ByVal intVal2)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines the parameter with the least value
	'*  Arguments supplied:		Look up
	'*  Return Value:			Value of parameter 1 or parameter 2, whichever is smaller
	'*  Called by:				Everyone
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		If (intVal1 > intVal2) Then
			Min = intVal2
		Else
			Min = intVal1
		End If
	End Function
	
	Public Function Max(ByVal intVal1, ByVal intVal2)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines the parameter with the greatest value
	'*  Arguments supplied:		Look up
	'*  Return Value:			Value of parameter 1 or parameter 2, whichever is larger
	'*  Called by:				Everyone
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		If (intVal1 > intVal2) Then
			Max = intVal1
		Else
			Max = intVal2
		End If

	End Function

	Public Function GetOSBits(ByVal objRemoteWMIServer, ByVal objRemoteRegServer, ByVal strConnectedWithThis, ByVal blnWMIGoodToGo, _
								ByVal blnWMIRegGoodToGo, ByVal blnRemoteRegGoodToGo, ByRef blnIs64BitMachine)	
	'*****************************************************************************************************************************************
	'*  Purpose:				Check if the processor address width is 64-bit
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Any
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim lngRegistryHive, strRegistryHive, strRegistryKey, strSearchValue, strSQLQuery, intErrNumber, strErrDescription, colWMI, objWMI
		Dim valRegValue, objShell, strOriginalKeyToSearch, strKeyToSearch, strToQuery, objExec, strStandardOut
		Const wbemFlagReturnWhenComplete = 0

		lngRegistryHive = m_HKEY_LOCAL_MACHINE
		strRegistryHive = "HKLM"
		strRegistryKey = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
		strSearchValue = "PROCESSOR_IDENTIFIER"
		'
		' Determine if the remote computer is a 32-bit OS or 64-bit OS
		'
		If (blnWMIGoodToGo) Then
			strSQLQuery = "SELECT AddressWidth FROM Win32_Processor WHERE AddressWidth=64"
			Call ExecWMI(objRemoteWMIServer, intErrNumber, strErrDescription, colWMI, strSQLQuery, wbemFlagReturnWhenComplete, Null)
			If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
				For Each objWMI In colWMI
					If (objWMI.AddressWidth = 64) Then
						blnIs64BitMachine = True
					Else
						blnIs64BitMachine = False
					End If
					Exit Function
				Next
			End If
		End If
		If (blnWMIRegGoodToGo) Then
			objRemoteRegServer.GetStringValue lngRegistryHive, strRegistryKey, strSearchValue, valRegValue
			If (Not IsNull(valRegValue)) Then
				If (InStr(valRegValue, "64") > 0) Then
					blnIs64BitMachine = True
				Else
					blnIs64BitMachine = False
				End If
				Exit Function
			End If
		End If
		If (blnRemoteRegGoodToGo) Then
			Set objShell = CreateObject("Wscript.Shell")
			strOriginalKeyToSearch = strRegistryHive & "\" & strRegistryKey
			strKeyToSearch = "\\" & strConnectedWithThis & "\" & strOriginalKeyToSearch
			If (InStr(strKeyToSearch, " ")) Then
				strKeyToSearch = Chr(34) & strKeyToSearch & Chr(34)
			End If
			If (InStr(strSearchValue, " ")) Then
				strToQuery = strKeyToSearch & " /v " & Chr(34) & strSearchValue & Chr(34)
			Else
				strToQuery = strKeyToSearch & " /v " & strSearchValue
			End If
			Set objExec = objShell.Exec("cmd /c REG QUERY " & strToQuery)
			strStandardOut = Trim(objExec.StdOut.ReadAll)
			'
			' Make sure the Reg Query was successful
			'
			If ((Not IsNull(strStandardOut)) And (Len(strStandardOut) <> 0) And _
				(InStr(1, strStandardOut, "ERROR: The system was unable to find the specified registry key or value.", vbTextCompare) = 0)) Then
				'
				' Make sure the original search string is in the output - indicates success
				'	HKEY_LOCAL_MACHINE\system\currentcontrolset\control\session manager\environment
				'	    PROCESSOR_IDENTIFIER    REG_SZ    Intel64 Family 6 Model 37 Stepping 5, GenuineIntel
				'
				If (InStr(1, strStandardOut, strSearchValue, vbTextCompare) > 0) Then
					If (InStr(strStandardOut, "64") > 0) Then
						blnIs64BitMachine = True
					Else
						blnIs64BitMachine = False
					End If
					'
					' Cleanup
					'
					Set objExec = Nothing
					Set objShell = Nothing
					Exit Function
				End If
			End If
		End If
		'
		' Cleanup
		'
		Set objExec = Nothing
		Set objShell = Nothing
		blnIs64BitMachine = False

	End Function

	Public Function DetermineNameOrIPAddress(ByVal strMachineNameOrIPAddress, ByRef blnFQDN, ByRef blnMachineName, ByRef blnIPAddress, _
													ByRef strMachineName, ByRef strIPAddress)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determine if the value passed is a NetBIOS name, FQDN, or IP Address
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strCheck, intCount, strHold, intLen, intCount1
	
		'
		' Initialize passed parameters
		'
		blnFQDN = False
		blnMachineName = False
		blnIPAddress = False
		strMachineName = ""
		strIPAddress = ""
		'
		' Determine how many periods there are in strWork
		'
		strCheck = Split(strMachineNameOrIPAddress, ".")
		'
		' Check for definite NetBIOS name
		'
		If (UBound(strCheck) = 0) Then
			'
			' Machine Name
			'
			blnMachineName = True
			strMachineName = strMachineNameOrIPAddress
			Exit Function
		End If
		'
		' Check for definite FQDN
		'
		If (UBound(strCheck) <> 3) Then
			'
			' Definitely an FQDN - strip off the Host name
			'
			blnFQDN = True
			strMachineName = strCheck(0)
			Exit Function
		End If
		'
		' Could be either a FQDN or IP Address (UBound = 3)
		'
		For intCount = 0 To UBound(strCheck)
			strHold = strCheck(intCount)
			intLen = Len(strHold)
			If (intLen > 3) Then
				'
				' More than 3 characters indicates a FQDN
				'
				blnFQDN = True
				strMachineName = strCheck(0)
				Exit Function
			End If
			'
			' Ensure that all characters are numeric
			'		
			For intCount1 = 1 To intLen
				If (IsNumeric(Mid(strHold, intCount1, 1)) = False) Then
					'
					' Non-numeric indicates a FQDN
					'
					blnFQDN = True
					strMachineName = strCheck(0)
					Exit Function
				End If
			Next
		Next
		'
		' If we are here then we have an IP Address
		'
		blnIPAddress = True
		strIPAddress = strMachineNameOrIPAddress
				
	End Function

	Public Function DetermineParameterType(ByVal strWork, ByRef blnIsDNSHostName, ByRef blnIsHostName, ByRef blnIsIPAddress)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determine if the value passed is a Host name, FQDN, or IP Address
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strCheck, intCount, strHold, intLen, intCount1
	
		blnIsDNSHostName = False
		blnIsHostName = False
		blnIsIPAddress = False
		'
		' Determine how many periods there are in strWork
		'
		If (InStr(strWork, ".") = 0) Then
			'
			' No periods - a NetBIOS name
			'
			blnIsHostName = True
			Exit Function
		End If
		'
		' We know there is a period in the value passed
		'
		strCheck = Split(strWork, ".")
		'
		' Check for definite FQDN
		'
		If (UBound(strCheck) <> 3) Then
			'
			' Definitely a DNSHostName - strip off the Host name
			'
			blnIsDNSHostName = True
			Exit Function
		End If
		'
		' Could be either a DNSHostName or IP Address (UBound = 3)
		'
		For intCount = 0 To UBound(strCheck)
			strHold = strCheck(intCount)
			intLen = Len(strHold)
			If (intLen > 3) Then
				'
				' More than 3 characters indicates a DNSHostName
				'
				blnIsDNSHostName = True
				Exit Function
			End If
			'
			' Ensure that all characters are numeric
			'		
			For intCount1 = 1 To intLen
				If (IsNumeric(Mid(strHold, intCount1, 1)) = False) Then
					'
					' Non-numeric indicates a DNSHostName
					'
					blnIsDNSHostName = True
					Exit Function
				End If
			Next
		Next
		'
		' If we are here then we have an IP Address
		'
		blnIsIPAddress = True
		
	End Function

	Public Function CreatePrintableString(ByRef strToConvert)
	'*****************************************************************************************************************************************
	'*  Purpose:				Convert the passed string to a printable one (replace unprintable characters with "(Axx)"
	'*  Arguments supplied:		Look up
	'*  Return Value:			Converted string
	'*  Called by:				LoadRS
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strTemp, intCount, strCharToConvert, intUCharacter, intCharacter
	
		'
		' Unicode: A character encoding standard developed by the Unicode Consortium. By using more than one byte to
		' represent each character, Unicode enables almost all of the written languages in the world to be represented
		' by using a single character set.
		'
		' Sometimes text values contain leading, trailing, or multiple embedded space characters (Unicode character set
		' values 32 and 160), or non-printing characters (Unicode character Set values 0 to 31, 127, 129, 141, 143, 144,
		' and 157). These characters can sometimes cause unexpected results when you sort, filter, or search.
		'
		' 7-bit ASCII, which is a subset of the ANSI character set (ANSI character set: An 8-bit character set used by
		' Microsoft Windows that allows you to represent up to 256 characters (0 through 255) by using your keyboard.
		' The ASCII character Set is a subset of the ANSI set. It's important to understand that the first 128 values
		' (0 to 127) in 7-bit ASCII represent the same characters as the first 128 values in the Unicode character set.
		'
		strTemp = ""
		For intCount = 1 To Len(strToConvert)
			strCharToConvert = Mid(strToConvert, intCount, 1)
			intUCharacter = AscW(strCharToConvert)
	 		intCharacter = Asc(strCharToConvert)
	 		If ((intUCharacter>=32 And intUCharacter<=126) Or (intUCharacter=9) Or (intUCharacter=10) Or (intUCharacter=13)) Then
				strTemp = strTemp & Chr(intCharacter)
			Else
				strTemp = strTemp & " Unprintable characters found"
				Exit For
			End If
		Next
		strToConvert = strTemp
	
	End Function

	Public Function DeleteAllRecordsetRows(ByRef rsToClear)
	'*****************************************************************************************************************************************
	'*  Purpose:				Delete all rows from the specified recordset
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Any
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		If (rsToClear.RecordCount > 0) Then
			If (Not rsToClear.BOF) Then
				rsToClear.MoveFirst
			End If
			While Not rsToClear.EOF
				rsToClear.Delete
				rsToClear.MoveNext
			Wend
		End If
	
	End Function

	Public Function GetRemoteOSVersion(ByVal intFlag, ByVal objRemoteWMIServer, ByVal objRemoteRegServer, ByVal strConnectedWithThis, _
										ByVal blnWMIGoodToGo, ByVal blnWMIRegGoodToGo, ByVal blnRemoteRegGoodToGo, ByRef strOSVersion, _
										ByRef objLogAndTrace)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get the information requested using either WMI or Registry calls
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				GatherMachineData
	'*  Calls:					LogThis, ExecWMI, GetRegistryEntry
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strSQLQuery, intErrNumber, strErrDescription, colWMI, objWMI, lngRegistryHive, strRegistryHive, strRegistryKey, strSearchValue
		Dim valRegValue, objShell, strKeyToSearch, strToQuery, objExec, strRead, strToParse, arrToParse
	
		strOSVersion = ""
		If (blnWMIGoodToGo) Then
			Call LogThis("Getting Version from Win32_OperatingSystem in GetRemoteOSVersion", objLogAndTrace)
			strSQLQuery = "SELECT Version FROM Win32_OperatingSystem"
			Call g_objFunctions.ExecWMI(objRemoteWMIServer, intErrNumber, strErrDescription, colWMI, strSQLQuery, intFlag, Null)
			If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
				For Each objWMI In colWMI
					strOSVersion = Split(objWMI.Version, ".", 3)(0) & "." & Split(objWMI.Version, ".", 3)(1)
					Call LogThis("Got OSVersion using Win32_OperatingSystem", objLogAndTrace)
					Exit Function
				Next
			End If
		End If
		lngRegistryHive = m_HKEY_LOCAL_MACHINE
		strRegistryHive = "HKLM"
		strRegistryKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
		strSearchValue = "CurrentVersion"
		If (blnWMIRegGoodToGo) Then
			objRemoteRegServer.GetStringValue lngRegistryHive, strRegistryKey, strSearchValue, valRegValue
			If (Not IsNull(valRegValue)) Then
				strOSVersion = valRegValue
				Exit Function
			End If
		End If
		If (blnRemoteRegGoodToGo) Then
			Set objShell = CreateObject("Wscript.Shell")
			strKeyToSearch = "\\" & strConnectedWithThis & "\" & strRegistryHive & "\" & strRegistryKey
			If (InStr(strKeyToSearch, " ")) Then
				strKeyToSearch = Chr(34) & strKeyToSearch & Chr(34)
			End If
			If (InStr(strSearchValue, " ")) Then
				strToQuery = strKeyToSearch & " /v " & Chr(34) & strSearchValue & Chr(34)
			Else
				strToQuery = strKeyToSearch & " /v " & strSearchValue
			End If
			Set objExec = objShell.Exec("cmd /c REG QUERY " & strToQuery)
			'
			' Make sure the original search string is in the output - indicates success
			'	HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion
			'		currentversion    REG_SZ    6.1
			'
			While Not objExec.StdOut.AtEndOfStream
				strRead = Trim(objExec.StdOut.ReadLine)
				If ((InStr(1, strRead, strSearchValue, vbTextCompare) > 0) And _
					(InStr(1, strRead, strRegistryKey, vbTextCompare) = 0)) Then
					'
					' Found what we were looking for...
					'
					strToParse = Trim(Split(strRead, strSearchValue, 2, vbTextCompare)(1))
					arrToParse = Split(strToParse, " ")
					strOSVersion = arrToParse(UBound(arrToParse))
					Set objExec = Nothing
					Exit Function
				End If
			Wend
		End If

	End Function

	Public Function GetOSVersion()
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines the Operating System version of the machine the script is executed on
	'*  Arguments supplied:		None
	'*  Return Value:			The verion of Windows
	'*  Called by:				ValidateOSVersion
	'*  Calls:					None
	'*	Requirements:			m_objWMILocal
	'*****************************************************************************************************************************************
		Dim objShell, strThisComputer, objWMIService, strSQLQuery, colWMI, intErrNumber, strErrDescription, objWMI, OSInfo, strMajorOSVersion
		Dim strMinorOSVersion, strBuildNumber, strWindowsVersion
		Const wbemFlagReturnWhenComplete = 0

		Set objShell = CreateObject("WScript.Shell")
		strThisComputer = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\" & strThisComputer & "\root\cimv2")
		'
		' Retrieve the OS version that is running on this machine ("5.0" = 2000; "5.1" = XP; "5.2" = 2003)
		'
		GetOSVersion = ""
		strSQLQuery = "SELECT Version FROM Win32_OperatingSystem"
		Call ExecWMI(objWMIService, intErrNumber, strErrDescription, colWMI, strSQLQuery, wbemFlagReturnWhenComplete, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
			For Each objWMI in colWMI
				'
				'  There should only be one colItem in the objWMI collection
				'
				OSInfo = Split(objWMI.Version, ".", 3)
				strMajorOSVersion = OSInfo(0)
				strMinorOSVersion = OSInfo(1)
				strBuildNumber = OSInfo(2)
				strWindowsVersion = Left(objWMI.Version, InStrRev(objWMI.Version, ".") - 1)
			Next
			GetOSVersion = strWindowsVersion
		End If
		'
		' Cleanup
		'
		Set colWMI = Nothing
		Set objWMIService = Nothing
		Set objShell = Nothing

	End Function

	Public Function ValidateOSVersion()
	'*****************************************************************************************************************************************
	'*  Purpose:				Ensures the machine running this script is at least XP or higher
	'*  Arguments supplied:		None
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					GetOSVersion
	'*  Requirements:			OSVersion Constants
	'*****************************************************************************************************************************************
		Dim strOSVersion
		
		strOSVersion = GetOSVersion()
		If (strOSVersion < m_OS_VERSION_XP) Then
			MsgBox "This program will only run on Windows XP or higher!", vbCritical + vbOKOnly, "Wrong Operating System" 
			WScript.Quit
		End If
	
	End Function

	Public Function VerifyAndLoad(ByVal varToVerify, ByVal intType)
	'*****************************************************************************************************************************************
	'*  Purpose:				Verifies the value passed and either loads a default value (based on type) or the actual value.
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					None
	'*  Requirements:			DEFAULT_DATE, VerifyAndLoad Constants
	'*****************************************************************************************************************************************
	Dim strHexHigh, strHexLow, bigintHigh, bigintLow, intHighAndLow, strYear, strMonth, strDay, strHour, strMinute, strSecond, arrDateTime
	Dim arrTime, strWorkMonth, strDateTime

	' Values for use with VarType function:
	'0 vbEmpty Empty (uninitialized) 
	'1 vbNull Null (no valid data) 
	'2 vbInteger Integer 
	'3 vbLong Long integer 
	'4 vbSingle Single-precision floating-point number 
	'5 vbDouble Double-precision floating-point number 
	'6 vbCurrency Currency 
	'7 vbDate Date 
	'8 vbString String 
	'9 vbObject Automation object 
	'10 vbError Error 
	'11 vbBoolean Boolean 
	'12 vbVariant Variant (used only with arrays of Variants) 
	'13 vbDataObject A data-access object 
	'17 vbByte Byte 
	'8192 vbArray Array (VBScript uses 8192 as a base for arrays and adds the code for the data type to indicate an array. 
	'					8204 indicates a variant array, the only real kind of array in VBScript.)

		Select Case VarType(varToVerify) 
			Case vbEmpty, vbNull
				Select Case intType
					Case vbBoolean
						VerifyAndLoad = False
					Case vbString
						VerifyAndLoad = ""
					Case vbInteger, vbLong, vbSingle, vbDouble
						VerifyAndLoad = -1
					Case vbDate, m_vbDateDNSTimestamp, m_vbDateMicrosoftDateTime, m_vbDateWMIDateTime, m_vbDateUTCDateTime
						VerifyAndLoad = m_DEFAULT_DATE
					Case Else
						VerifyAndLoad = Null
				End Select
			Case vbInteger, vbLong
				If (intType = m_vbDateDNSTimestamp) Then
					'
					' This is an unformatted DateTime string
					'
					If (varToVerify = 0) Then
						VerifyAndLoad = m_MSSQL_TIMESTAMP_DATE_BASE
					Else
						VerifyAndLoad = DateAdd("H", varToVerify, m_DNS_TIMESTAMP_DATE_BASE)
					End If
				ElseIf (IsNumeric(varToVerify)) Then
					If ((varToVerify >= -32768) And (varToVerify <= 32767)) Then
						'
						' Valid value range for Integer is -32,768 to 32,767
						'
						VerifyAndLoad = CInt(varToVerify)
					ElseIf ((varToVerify >= -2147483648) And (varToVerify <= 2147483647)) Then
						'
						' Valid value range for Long is -2,147,483,648 to 2,147,483,647
						'
						VerifyAndLoad = CLng(varToVerify)
					End If
				Else
					VerifyAndLoad = -1
				End If
			Case vbSingle, vbDouble
				If (IsNumeric(varToVerify)) Then
					If (intType = vbSingle) Then
						'
						' Valid value range for Single is -3.402823E38 to -1.401298E-45 and
						'									1.401298E-45 to 3.402823E38
						'
						VerifyAndLoad = CSng(varToVerify)
					Else
						'
						' Valid value range for Double is -1.79769313486232E308 To -4.94065645841247E-324 and
						'									4.94065645841247E-324 To 1.79769313486232E308
						'
						VerifyAndLoad = CDbl(varToVerify)
					End If
				Else
					VerifyAndLoad = -1.0
				End If
'			Case vbCurrency
'				Call MsgBox("Currency")
			Case vbDate
				If (intType = m_vbDateDNSTimestamp) Then
					'
					' This is an unformatted DateTime string
					'
					If (varToVerify = 0) Then
						VerifyAndLoad = m_MSSQL_TIMESTAMP_DATE_BASE
					Else
						VerifyAndLoad = DateAdd("H", varToVerify, m_DNS_TIMESTAMP_DATE_BASE)
					End If
				ElseIf (IsDate(varToVerify)) Then
					VerifyAndLoad = FormatDateTime(varToVerify)
				Else
					VerifyAndLoad = m_DEFAULT_DATE
				End If
			Case vbString
' 				WScript.Echo varToVerify & vbTab & intType
				If (intType = m_vbDateMicrosoftDateTime) Then
' 					WScript.Echo "MicrosoftDateTime type"
					'
					' Date is in the following format:
					'	Byte 0: Year = 1970 + valRegValue[0] 
					'	Byte 1: Month = valRegValue[1] + 1 (zero based)
					'	Byte 2: Day = valRegValue[2]
					'	Byte 3: Hour = valRegValue[3]
					'	Byte 4: Minute = valRegValue[4]
					'	Byte 5: Second = valRegValue[5]
					'	Byte 6: HundredSecond = valRegValue[6]
					'	Byte 7: ThousandSecond = valRegValue[7]
					'
					strYear = HexToDec(Mid(varToVerify, 1, 2)) + 1970
					strMonth = HexToDec(Mid(varToVerify, 3, 2)) + 1
					strDay = HexToDec(Mid(varToVerify, 5, 2))
					strHour = HexToDec(Mid(varToVerify, 7, 2))
					strMinute = HexToDec(Mid(varToVerify, 9, 2))
					strSecond = HexToDec(Mid(varToVerify, 11, 2))
					VerifyAndLoad = MassageTimestamp(CDate(strMonth & "/" & strDay & "/" & strYear & " " & strHour & ":" & strMinute & ":" & strSecond))
				ElseIf (intType = m_vbDateUTCDateTime) Then
					'
					' Here is the format we've seen in the field: Wed Jul 10 15:09:30 UTC 2013
					'
					arrDateTime = Split(varToVerify, " ")
					arrTime = Split(arrDateTime(3), ":")
					strYear = arrDateTime(5)
					strWorkMonth = UCase(Mid(arrDateTime(1), 1, 3))
					Select Case strWorkMonth
						Case "JAN"
							strMonth = "01"
						Case "FEB"
							strMonth = "02"
						Case "MAR"
							strMonth = "03"
						Case "APR"
							strMonth = "04"
						Case "MAY"
							strMonth = "05"
						Case "JUN"
							strMonth = "06"
						Case "JUL"
							strMonth = "07"
						Case "AUG"
							strMonth = "08"
						Case "SEP"
							strMonth = "09"
						Case "OCT"
							strMonth = "10"
						Case "NOV"
							strMonth = "11"
						Case "DEC"
							strMonth = "12"
						Case Else
					End Select
					strDay = arrDateTime(2)
					strHour = arrTime(0)
					If (Len(strHour) < 2) Then
						strHour = "0" & strHour
					End If
					strMinute = arrTime(1)
					If (Len(strMinute) < 2) Then
						strMinute = "0" & strMinute
					End If
					strSecond = arrTime(2)
					If (Len(strSecond) < 2) Then
						strSecond = "0" & strSecond
					End If
					strDateTime = strYear & strMonth & strDay & strHour & strMinute & strSecond
					VerifyAndLoad = CDate(WMIDateStringToDate(strDateTime))
				ElseIf (intType = m_vbDateWMIDateTime) Then
					VerifyAndLoad = CDate(WMIDateStringToDate(varToVerify))
				ElseIf (intType = vbDate) Then
					If (varToVerify = "") Then
						VerifyAndLoad = m_DEFAULT_DATE
					ElseIf (IsNumeric(Mid(varToVerify,1,2)) And IsAlpha(Mid(varToVerify,3,1)) And Len(varToVerify) = 16) Then
						strHexHigh = Mid(varToVerify,1,8)
						strHexLow = Mid(varToVerify,9,8)
						bigintHigh = HexToDec(strHexHigh)
						bigintLow = HexToDec(strHexLow)
						intHighAndLow = bigintHigh * (2^32) + bigintLow
						intHighAndLow = intHighAndLow / (60 * 10000000)
						intHighAndLow = intHighAndLow / 1440 
						VerifyAndLoad = intHighAndLow + #1/1/1601#
					ElseIf (IsNumeric(Mid(varToVerify,1,6))) Then
						'
						' This is an unformatted DateTime string
						'
						VerifyAndLoad = CDate(WMIDateStringToDate(varToVerify))
					ElseIf ((InStr(varToVerify, "/") > 0) Or (InStr(varToVerify, ".") > 0)) Then
						If (InStr(varToVerify, ".") > 0) Then
							VerifyAndLoad = FormatDateTime(Replace(varToVerify, ".", "/"))
						Else
							VerifyAndLoad = FormatDateTime(varToVerify)
						End If
					Else
						VerifyAndLoad = varToVerify
					End If
				ElseIf ((intType = vbInteger) Or (intType = vbLong) Or (intType = vbDouble)) Then
					If ((varToVerify = "") Or (varToVerify = " ")) Then
						VerifyAndLoad = -1
					Else
						If ((varToVerify >= -32768) And (varToVerify <= 32767)) Then
							'
							' Valid value range for Integer is -32,768 to 32,767
							'
							VerifyAndLoad = CInt(varToVerify)
						ElseIf ((varToVerify >= -2147483648) And (varToVerify <= 2147483647)) Then
							'
							' Valid value range for Long is -2,147,483,648 to 2,147,483,647
							'
							VerifyAndLoad = CLng(varToVerify)
						Else
							'
							' Valid value range for Double is -1.79769313486232E308 To -4.94065645841247E-324 and
							'									4.94065645841247E-324 To 1.79769313486232E308
							'
							VerifyAndLoad = CDbl(varToVerify)
						End If
					End If
				Else
					VerifyAndLoad = RemoveControlCharacters(varToVerify)
				End If
			Case vbObject
				Set VerifyAndLoad = varToVerify
'			Case vbError
'				Call MsgBox("Error")
			Case vbBoolean
				If (varToVerify = False) Then
					VerifyAndLoad = False
				Else
					VerifyAndLoad = True
				End If
'			Case vbDataObject
'				Call MsgBox("DataObject")
'			Case vbByte
'				Call MsgBox("Byte")
			Case vbArray + vbVariant
				'
				' Multi-valued strings in AD will be in this case
				'
				If (intType = vbString) Then
					'
					' This is an Array to be loaded into a string field
					'
					VerifyAndLoad = RemoveControlCharacters(Join(varToVerify, ";;"))
				ElseIf (intType = vbInteger) Then
					'
					' This is an Integer Array
					'
					VerifyAndLoad = RemoveControlCharacters(Join(varToVerify, ";;"))
				Else
					VerifyAndLoad = varToVerify
				End If
			Case vbArray + vbByte
				'
				' Octet String or SID data types from AD will be in this case
				'
				If (intType = m_vbGUID) Then
					'
					' This is a GUID to be loaded into a string field
					'
					VerifyAndLoad = ConvertObjectGUIDToStringGUID(varToVerify)
				ElseIf (intType = m_vbSID) Then
					'
					' This is a SID to be loaded into a string field
					'
					VerifyAndLoad = ConvertObjectSIDToStringSID(varToVerify)
				ElseIf (intType = m_vbScheduleOrRelay) Then
					'
					' This is an Activation Schedule or Relay IP List to be loaded into a string field
					'
					VerifyAndLoad = OctetToHexStr(varToVerify)
				Else
					VerifyAndLoad = varToVerify
				End If
			Case Else
				VerifyAndLoad = varToVerify
		End Select

	End Function

	Public Function GetMySID(ByVal objWMIService, ByRef strSID)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determine our process SID
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 if successful, -1 if unsuccessful.
	'*  Called by:				All
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strCommand, objStartup, objConfig, objProcess, intFlag, strSQLQuery, intErrNumber, strErrDescription, intRetVal, intPID, colProcesses
		Dim strOwnerSID
	
		Const NORMAL_WINDOW = 1
		Const HIDDEN_WINDOW = 12
		Const wbemFlagReturnWhenComplete = 0
	
		strSID = Empty
		strCommand = "Wordpad.exe"
		Set objStartup = objWMIService.Get("Win32_ProcessStartup")
		Set objConfig = objStartup.SpawnInstance_
		objConfig.ShowWindow = HIDDEN_WINDOW
		'
		' Create a new process.
		'
		On Error Resume Next
		Set objProcess = objWMIService.Get("Win32_Process")
		intErrNumber = Err.Number
		On Error GoTo 0
		
		If (intErrNumber = 0) Then
			intRetVal = objProcess.Create(strCommand, Null, objConfig, intPID)
			If (intRetVal = 0) Then
				intFlag = 48
				strSQLQuery = "SELECT * FROM Win32_Process WHERE ProcessID='" & intPID & "'"
				Call ExecWMI(objWMIService, intErrNumber, strErrDescription, colProcesses, strSQLQuery, wbemFlagReturnWhenComplete, Null)
				If ((intErrNumber=0) And (UCase(TypeName(colProcesses)) = "SWBEMOBJECTSET")) Then
					For Each objProcess In colProcesses
						intRetVal = objProcess.GetOwnerSID(strOwnerSID)
						If (intRetVal = 0) Then
							strSID = UCase(strOwnerSID)
						End If
						On Error Resume Next
						objProcess.Terminate()
						On Error GoTo 0
					Next
				End If
			End If
		End If
		'
		' Cleanup
		'
		Set objStartup = Nothing
		Set objConfig = Nothing
		Set objProcess = Nothing
		Set colProcesses = Nothing
	
	End Function

	Public Function GetMyProcessID(ByVal objWMIService, ByVal strScriptOwnerSID)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines my PID
	'*  Arguments supplied:		None
	'*  Return Value:			My PID (integer value)
	'*  Called by:				Any
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strScriptFullName, strScriptName, intFlag, strSQLQuery, colProcesses, intErrNumber, strErrDescription, objProcess, strCmdLine
		Dim intPID, intRetVal, strOwnerSID
		Const wbemFlagForwardOnly = 32
		Const wbemFlagReturnImmediately = 16
		
		strScriptFullName = WScript.ScriptFullName
		strScriptName = WScript.ScriptName

		On Error Resume Next
		intFlag = wbemFlagForwardOnly + wbemFlagReturnImmediately
		strSQLQuery = "SELECT * FROM Win32_Process WHERE Name='cscript.exe'"
		Call ExecWMI(objWMIService, intErrNumber, strErrDescription, colProcesses, strSQLQuery, intFlag, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colProcesses)) = "SWBEMOBJECTSET")) Then
			For Each objProcess In colProcesses
				strCmdLine = objProcess.commandLine
				intPID = objProcess.processId
				On Error Resume Next
				intRetVal = objProcess.GetOwnerSID(strOwnerSID)
				intErrNumber = Err.Number
				On Error GoTo 0
				If ((intRetVal = 0) And (intErrNumber = 0)) Then
					If (InStr(1, strCmdLine, strScriptFullName, vbTextCompare) > 0) Then
						If (IsEmpty(strScriptOwnerSID)) Then
							GetMyProcessID = intPID
							Set colProcesses = Nothing
							Exit Function
						Else
							If (InStr(1, strScriptOwnerSID, strOwnerSID, vbTextCompare) > 0) Then
								GetMyProcessID = intPID
								Set colProcesses = Nothing
								Exit Function
							End If
						End If
					End If
				End If
			Next
		End If
		GetMyProcessID = intPID
		'
		' Cleanup
		'
		Set colProcesses = Nothing
	
	End Function

	Public Function GetOwnerInfo(ByVal objWMIService, ByVal strThisComputer, ByVal intPID, ByRef blnIsLocalUser)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines the owner of the running process (and whether it is local or domain)
	'*  Arguments supplied:		intPID (current process ID)
	'*  Return Value:			strOwnerName (owner of process)
	'*							strOwnerDomain (either local machine or domain name)
	'*							blnIsLocalUser (true if local user account, false if domain account)
	'*  Called by:				Any
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strSQLQuery, colWMI, intErrNumber, strErrDescription, objWMI, intRetVal, strOwnerName, strOwnerDomain
		Const wbemFlagReturnWhenComplete = 0
	
		blnIsLocalUser = False
		strSQLQuery = "SELECT * FROM Win32_Process WHERE ProcessID = " & intPID
		Call ExecWMI(objWMIService, intErrNumber, strErrDescription, colWMI, strSQLQuery, wbemFlagReturnWhenComplete, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
			For Each objWMI In colWMI
				intRetVal = objWMI.GetOwner(strOwnerName, strOwnerDomain)
				If (intRetVal <> 0) Then
					'
					' This should never happen
					'
					strOwnerName = Null
					strOwnerDomain = Null
					blnIsLocalUser = False
					GetOwnerInfo = -1
					Exit Function
				Else 
					'
					' The strOwnerDomain variable will contain one of two possible values:
					'	The name of the local computer (if being executed by a local user)
					'	The name of a domain (if being executed by a domain user)
					'
					If (InStr(1, strThisComputer, strOwnerDomain, vbTextCompare) > 0) Then
						'
						' Found our computer name in the strOwnerDomain variable...local user
						'
						blnIsLocalUser = True
					Else
						blnIsLocalUser = False
					End If
				End If
			Next
		End If
		'
		' Cleanup
		'
		Set colWMI = Nothing
		GetOwnerInfo = 0
	
	End Function

	Public Function GetMyPID(ByRef intPID)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get the Process ID (PID) from this process
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim objShell, strThisComputer, objWMIService, strScriptOwnerSID

		Set objShell = CreateObject("WScript.Shell")
		strThisComputer = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\" & strThisComputer & "\root\cimv2")
		'
		' Who owns the running process and is it being executed with a local or domain account
		'
		Call GetMySID(objWMIService, strScriptOwnerSID)
		intPID = GetMyProcessID(objWMIService, strScriptOwnerSID)
		'
		' Cleanup
		'
		Set objWMIService = Nothing
		Set objShell = Nothing

	End Function

	Public Function GetTimeZoneBias(ByRef lngAdjust)
	'*****************************************************************************************************************************************
	'*  Purpose:				Gets the TimeZone Bias (offset) from the registry
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					None
	'*****************************************************************************************************************************************
		Dim objShell, lngBiasKey, intCount
		
		Set objShell = CreateObject("WScript.Shell")
		lngAdjust = 0
		'
		' Obtain local Time Zone bias from machine registry.
		'
		lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
		If (UCase(TypeName(lngBiasKey)) = "LONG") Then
			lngAdjust = lngBiasKey
		ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
			lngAdjust = 0
			For intCount = 0 To UBound(lngBiasKey)
				lngAdjust = lngAdjust + (lngBiasKey(intCount) * 256^k)
			Next
		End If

	End Function

	Public Function Integer8Date(ByVal objDate, ByVal intAdjustmentType)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts an Integer8Date (64-bit) value to a date
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					None
	'*****************************************************************************************************************************************
	'
	' Many attributes in Active Directory have a syntax called Integer8.  These 64-bit numbers (8 bytes) usually represent time In
	' 100-nanosecond intervals.  If the Integer8 attribute is a date, the value represents the number of 100-nanosecond intervals
	' since 12:00 AM January 1, 1601.  ADSI automatically employs the IADsLargeInteger interface to deal with these 64-bit numbers.
	' This interface has two property methods, HighPart and LowPart, which break the number up into two 32-bit numbers.  The LowPart
	' property method returns values between -2^31 and 2^31 - 1.  Whenever the LowPart method returns a negative value, the calculation
	' is wrong by 7 minutes, 9.5 seconds (2^32 100-nanosecond intervals).  The work-around is to increase the value returned by the
	' HighPart method by one whenever the value returned by the LowPart method is negative.  The code below gives correct results in all cases.
	'
		Dim lngHigh, lngLow, lngBiasKey, lngAdjust, intCount, lngDate, intErrNumber

		lngHigh = objDate.HighPart
		lngLow = objDate.LowPart

		If ((lngHigh = -1) And (lngLow = -1)) Then
			Integer8Date = m_DEFAULT_DATE
		Else
			lngAdjust = 0
			If ((intAdjustmentType = m_INTEGER8_DATE_ADJUST_LOCAL) Or (intAdjustmentType = m_INTEGER8_DATE_ADJUST_GMT)) Then
				Call GetTimeZoneBias(lngAdjust)
				If (intAdjustmentType = m_INTEGER8_DATE_ADJUST_GMT) Then
					lngAdjust = lngAdjust * -1
				End If
			End If
			'
			' Account for error in IADslargeInteger property methods.
			'
			If (lngLow < 0) Then
				lngHigh = lngHigh + 1
			End If
			If ((lngHigh = 0) And (lngLow = 0)) Then
				lngAdjust = 0
			End If
			'
			' The default date for Microsoft date fields uses 1/1/1601 as the base value - don't change
			'
			lngDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) + lngLow) / 600000000 - lngAdjust) / 1440
			'
			' Trap error if lngDate is ridiculously huge.
			'
			On Error Resume Next
			Integer8Date = CDate(lngDate)
			intErrNumber = Err.Number
			On Error GoTo 0
			
			If ((intErrNumber <> 0) Or (lngDate = "1/1/1601")) Then
				Integer8Date = m_DEFAULT_DATE
			End If
		End If

	End Function

	Public Function Integer8Time(ByVal objDate, ByVal intConversionType)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts an Integer8 (64-bit) value to a number representing seconds, minutes, hours, or days
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					None
	'*  Requirements:			Integer8 Constants
	'*****************************************************************************************************************************************
	'
	' Many attributes in Active Directory have a syntax called Integer8.  These 64-bit numbers (8 bytes) usually represent time In
	' 100-nanosecond intervals.  If the Integer8 attribute is a date, the value represents the number of 100-nanosecond intervals
	' since 12:00 AM January 1, 1601.  ADSI automatically employs the IADsLargeInteger interface to deal with these 64-bit numbers.
	' This interface has two property methods, HighPart and LowPart, which break the number up into two 32-bit numbers.  The LowPart
	' property method returns values between -2^31 and 2^31 - 1.  Whenever the LowPart method returns a negative value, the calculation
	' is wrong by 7 minutes, 9.5 seconds (2^32 100-nanosecond intervals).  The work-around is to increase the value returned by the
	' HighPart method by one whenever the value returned by the LowPart method is negative.  The code below gives correct results in all cases.
	'
		Dim lngHigh, lngLow, lngDuration
	
		Const MILLISECS_IN_SECOND		= 10000000
		Const SECONDS_IN_MINUTE			= 60
		Const MINUTES_IN_HOUR			= 60
		Const HOURS_IN_DAY				= 24

		Const INTEGER8_TIME_SECONDS		= 1
		Const INTEGER8_TIME_MINUTES		= 2
		Const INTEGER8_TIME_HOURS		= 3
		Const INTEGER8_TIME_DAYS		= 4
	
		lngHigh = objDate.HighPart
		lngLow = objDate.LowPart
		'
		' Account for error in IADslargeInteger property methods.
		'
		If (lngLow < 0) Then
			lngHigh = lngHigh + 1
		End If
		lngDuration = lngHigh * (2 ^ 32) + lngLow
		Select Case intConversionType
			Case INTEGER8_TIME_SECONDS
				Integer8Time = -lngDuration / (MILLISECS_IN_SECOND)
			Case INTEGER8_TIME_MINUTES
				Integer8Time = -lngDuration / (MILLISECS_IN_SECOND * SECONDS_IN_MINUTE)
			Case INTEGER8_TIME_HOURS
				Integer8Time = -lngDuration / (MILLISECS_IN_SECOND * SECONDS_IN_MINUTE * MINUTES_IN_HOUR)
			Case INTEGER8_TIME_DAYS
				Integer8Time = -lngDuration / (MILLISECS_IN_SECOND * SECONDS_IN_MINUTE * MINUTES_IN_HOUR * HOURS_IN_DAY)
			Case Else
		End Select
			
	End Function

	Public Function BinToDec(ByVal strBin)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts a binary string to a decimal value
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim lngResult, intIndex, strDigit
		
		lngResult = 0
		For intIndex = Len(strBin) To 1 Step -1
			strDigit = Mid(strBin, intIndex, 1)
			Select Case strDigit
				Case "0"
					'
					' Do nothing
					'
				Case "1"
					lngResult = lngResult + (2 ^ (Len(strBin) - intIndex))
				Case Else
					'
					' Invalid Binary digit, so the whole thing is invalid
					'
					lngResult = 0
					intIndex = 0
			End Select
		Next
		BinToDec = lngResult
	
	End Function

	Public Function DecToBin(ByVal DecValue)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts a decimal value to a binary string
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strResult, TempValue, BinValue
	
		strResult = ""
		While DecValue > 0
			TempValue = DecValue Mod 2
			BinValue = CStr(TempValue) + BinValue
			DecValue = DecValue \ 2
		Wend
		DecToBin = BinValue
		
	End Function

	Public Function GetParentFolder()
	'*****************************************************************************************************************************************
	'*  Purpose:				Gets the parent folder based on this scripts' location
	'*  Arguments supplied:		Look up
	'*  Return Value:			Parent folder path
	'*  Called by:				All
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim objFSO, strParentFolder
	
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		strParentFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
		If (Right(strParentFolder, 1) <> "\") Then
			strParentFolder = strParentFolder & "\"
		End If
		GetParentFolder = strParentFolder
		'
		' Cleanup
		'
		Set objFSO = Nothing

	End Function

	Public Function BuildPath(ByVal strFolderName)
	'*****************************************************************************************************************************************
	'*  Purpose:				Create all folders required to build a given path.
	'*  Arguments supplied:		strFolderName
	'*  Return Value:			0 to indicate success
	'*  Called by:				none
	'*  Calls:					none
	'*  Requirements:			ADO Constants
	'*****************************************************************************************************************************************
	' ---------------------------------------------------------------
	' Start at the path given and check to see if it exists. If not
	' then walk the up-level path (using the PARENTFOLDER method)
	' adding each unfound folder path to an array.  When the path is
	' found where a folder exists the array is "walked" and directories
	' created that didn't exist.  This function works with UNC's as
	' long as the share exists.
	' ---------------------------------------------------------------
		Dim objFSO, rsGeneric, strFolderPath, blnISaySo
		
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set rsGeneric = CreateObject("ADODB.Recordset")
		rsGeneric.Fields.Append "SavedData", m_adVarChar, 255
		rsGeneric.Open
		strFolderPath = strFolderName
		blnISaySo = True
		While blnISaySo
			If (Not objFSO.FolderExists(strFolderPath)) Then
				'
				' If the folder doesn't exist then add it to the array
				' preserving the data that is already in the array.
				'
				rsGeneric.AddNew
				rsGeneric("SavedData") = strFolderPath
				rsGeneric.Update
				strFolderPath = objFSO.GetParentFolderName(strFolderPath)
			Else
				'
				' We won't get here until the subroutine gets to a directory
				' that already exists.
				'
				If (rsGeneric.RecordCount > 0) Then
					rsGeneric.MoveLast
					While Not rsGeneric.BOF
						'
						' Loop through the array backwards creating any non-existent folders.
						'
						If (Not objFSO.FolderExists(rsGeneric("SavedData"))) Then
							On Error Resume Next
							objFSO.CreateFolder(rsGeneric("SavedData"))
							On Error GoTo 0
						End If
						rsGeneric.MovePrevious
					Wend
				End If
				blnISaySo = False
			End If
		Wend
		rsGeneric.Close
		Set rsGeneric = Nothing
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		BuildPath = 0

	End Function

	Public Function IsNamespaceInstalled(ByVal strConnectedWithThis, ByVal strNamespaceForConnection, ByVal strNamespaceToValidate, _
											ByVal strUserID, ByVal strPassword, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if the specified Namespace is installed
	'*  Arguments supplied:		Look up
	'*  Return Value:			True to indicate Namespace installed, False to indicate Namespace not installed
	'*  Called by:				LoadExchangeMailboxStore, GatherMachineData
	'*  Calls:					CreateServerConnection
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim objConnection, intRetVal, objWMIService, intErrNumber, strErrDescription, colNameSpace, objNameSpace, strError

		intRetVal = CreateServerConnection(strConnectedWithThis, objWMIService, intErrNumber, strErrDescription, strError, _
											strNamespaceForConnection, strUserID, strPassword, objLogAndTraceErrors)
		If (intRetVal = 0) Then
			Set colNameSpace = objWMIService.Instancesof("__NameSpace")
			For Each objNameSpace In colNameSpace
				If (InStr(1, objNameSpace.Name, strNamespaceToValidate, vbTextCompare) > 0) Then
					'
					' Cleanup
					'
					Set objWMIService = Nothing
					Set objConnection = Nothing
					IsNamespaceInstalled = True
					Exit Function
				End If
			Next
		End If
		'
		' Cleanup
		'
		Set objWMIService = Nothing
		Set objConnection = Nothing
		IsNamespaceInstalled = False
	
	End Function

	Public Function CreateCurrentTimestamp()
	'*****************************************************************************************************************************************
	'*  Purpose:				Creates a date/time string in the following format: "mm/dd/yyyy hh:mm:ss xm"
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Main()
	'*  Calls:					None
	'*  Requirements:			ADO Constants, WBEM Constants
	'*****************************************************************************************************************************************
		Dim dtCurrent, intDay, intMonth, intYear, tmCurrent, intHour, intMinute, intSecond, strMonth, strDay, strYear, strAMPM, strHour
		Dim strMinute, strSecond

		'
		' Format the Date
		'
		dtCurrent = Date()
		intDay = Day(dtCurrent)
		intMonth = Month(dtCurrent)
		intYear = Year(dtCurrent)
		'
		' Format the Time
		'
		tmCurrent = Time()
		intHour = Hour(tmCurrent)
		intMinute = Minute(tmCurrent)
		intSecond = Second(tmCurrent)
		'
		' Format Month
		'
		If (intMonth < 10) Then
			strMonth = "0" & CStr(intMonth)
		Else
			strMonth = CStr(intMonth)
		End If
		'
		' Format Day
		'
		If (intDay < 10) Then
			strDay = "0" & CStr(intDay)
		Else
			strDay = CStr(intDay)
		End If
		'
		' Format Year
		'
		strYear = CStr(intYear)
		'
		' Determine if the Time (in 24-hour format) is AM or PM
		'
		If ((intHour >= 0) And (intHour <= 11)) Then
			strAMPM = "AM"
		Else
			strAMPM = "PM"
		End If
		'
		' Convert hour to standard time vs. military time
		'
		If (intHour > 12) Then
			intHour = intHour - 12
		End If
		'
		' Adjust any Time values that are 0 to contain 2 digits
		'
		If (intHour < 10) Then
			strHour = "0" & CStr(intHour)
		Else
			strHour = CStr(intHour)
		End If
		If (intMinute < 10) Then
			strMinute = "0" & CStr(intMinute)
		Else
			strMinute = CStr(intMinute)
		End If
		If (intSecond < 10) Then
			strSecond = "0" & CStr(intSecond)
		Else
			strSecond = CStr(intSecond)
		End If
		CreateCurrentTimestamp = strMonth & "/" & strDay & "/" & strYear & " " & strHour & ":" & strMinute & ":" & strSecond & " " & strAMPM

	End Function

	Public Function BuildDateString(ByVal dtmDate)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts a WMI Date string to a valid Date
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Main()
	'*  Calls:					None
	'*  Requirements:			ADO Constants, WBEM Constants
	'*****************************************************************************************************************************************
		Dim arrDateTime, strDate, strTime, strAMPM, intMonth, intDay, intYear, intHour, intMinute, intSecond, strMonth, strYear, strDay
		Dim strHour, strMinute, strSecond, strOutputTime

		'
		' I've now seen 2 different date/time strings: "5/21/2014 16:34:39" and "5/21/2014 11:35:53 AM"
		' Code for both below...
		'
		arrDateTime = Split(dtmDate, " ")
		'
		' Get Date information
		'
		strDate = arrDateTime(0)
		intMonth = CInt(Split(strDate, "/")(0))
		intDay = CInt(Split(strDate, "/")(1))
		intYear = CInt(Split(strDate, "/")(2))
		'
		' Get Time information
		'
		strTime = arrDateTime(1)
		intHour = CInt(Split(strTime, ":")(0))
		intMinute = CInt(Split(strTime, ":")(1))
		intSecond = CInt(Split(strTime, ":")(2))
		'
		' Here's where the difference will occur
		'
		If (UBound(arrDateTime) = 2) Then
			strAMPM = arrDateTime(2)
		Else
			If (intHour >= 12) Then
				strAMPM = "PM"
			Else
				strAMPM = "AM"
			End If
		End If
		'
		' Format Month
		'
		If (intMonth < 10) Then
			strMonth = "0" & CStr(intMonth)
		Else
			strMonth = CStr(intMonth)
		End If
		'
		' Format Day
		'
		If (intDay < 10) Then
			strDay = "0" & CStr(intDay)
		Else
			strDay = CStr(intDay)
		End If
		'
		' Format Year
		'
		strYear = CStr(intYear)
		'
		' Format Hour
		'
		If (UCase(strAMPM) = "PM") Then
			If (intHour <> 12) Then
				strHour = CStr(intHour + 12)
			Else
				strHour = CStr(intHour)
			End If
		Else
			If (intHour = 12) Then
				strHour = "00"
			ElseIf (intHour < 10) Then
				strHour = "0" & CStr(intHour)
			Else
				strHour = CStr(intHour)
			End If
		End If
		'
		' Format Minute
		'
		If (intMinute < 10) Then
			strMinute = "0" & CStr(intMinute)
		Else
			strMinute = CStr(intMinute)
		End If
		'
		' Format Second
		'
		If (intSecond < 10) Then
			strSecond = "0" & CStr(intSecond)
		Else
			strSecond = CStr(intSecond)
		End If
		strOutputTime = strYear & strMonth & strDay & strHour & strMinute & strSecond
		BuildDateString = strOutputTime

	End Function

	Public Function WMIDateStringToDate(ByVal dtmDate)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts a WMI Date string to a valid Date
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Main()
	'*  Calls:					None
	'*  Requirements:			ADO Constants
	'*****************************************************************************************************************************************
		Dim strYear, strMonth, strDay, intHour, intMinute, intSecond, strAMPM, strHour, strMinute, strSecond, strOutputTime
	
		If ((IsNull(dtmDate)) Or (Len(dtmDate) = 0)) Then
			WMIDateStringToDate = m_DEFAULT_DATE
			Exit Function
		End If
		'
		' Call CInt to remove leading zeros then CStr to load to a string variable
		'
		strYear = CStr(CInt(Left(dtmDate, 4)))
		strMonth = CStr(CInt(Mid(dtmDate, 5, 2)))
		strDay = CStr(CInt(Mid(dtmDate, 7, 2)))
		If (Len(dtmDate) = 8) Then
			'
			' No time included
			'
			strHour = "00"
			strMinute = "00"
			strSecond = "00"
			strAMPM = ""
		Else
			'
			' Call CInt to remove leading zeros but save the Time fields in integer form for manipulation
			'
			intHour = CInt(Mid(dtmDate, 9, 2))
			intMinute = CInt(Mid(dtmDate, 11, 2))
			intSecond = CInt(Mid(dtmDate,13, 2))
			'
			' Determine if the Time (in 24-hour format) is AM or PM
			'
			If ((intHour >= 0) And (intHour <= 11)) Then
				strAMPM = "AM"
			Else
				strAMPM = "PM"
			End If
			'
			' Convert hour to standard time vs. military time
			'
			If (intHour > 12) Then
				intHour = intHour - 12
			End If
			'
			' Adjust any Time values that are 0 to contain 2 digits
			'
			If (intHour = 0) Then
				strHour = "00"
			Else
				strHour = CStr(intHour)
			End If
			If (intMinute < 10) Then
				strMinute = "0" & CStr(intMinute)
			Else
				strMinute = CStr(intMinute)
			End If
			If (intSecond < 10) Then
				strSecond = "0" & CStr(intSecond)
			Else
				strSecond = CStr(intSecond)
			End If
		End If
		strOutputTime = strMonth & "/" & strDay & "/" & strYear & " " & strHour & ":" & strMinute & ":" & strSecond & " " & strAMPM
		WMIDateStringToDate = strOutputTime
	
	End Function

	Public Function WMIFileExists(ByRef objRemoteWMIServer, ByVal strFilePath, ByVal intFlag)
	'*****************************************************************************************************************************************
	'*  Purpose:				Checks to see if a file exists (using CIM_DataFile WMI call).
	'*  Arguments supplied:		Look up
	'*  Return Value:			-1 to indicate file exists; 0 to indicate file doesn't exist
	'*  Called by:				Main()
	'*  Calls:					None
	'*  Requirements:			ADO Constants
	'*****************************************************************************************************************************************
		Dim strSQLQuery, intErrNumber, strErrDescription, colFiles

		WMIFileExists = False
		strSQLQuery = "SELECT Name FROM CIM_DataFile WHERE Name = '" & Replace(strFilePath, "\", "\\") & "'"
		Call ExecWMI(objRemoteWMIServer, intErrNumber, strErrDescription, colFiles, strSQLQuery, intFlag, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colFiles))="SWBEMOBJECTSET")) Then
			If (colFiles.Count = 0) Then
				WMIFileExists = False
			Else
				WMIFileExists = True
			End If
		End If
		Set colFiles = Nothing
	
	End Function

	Public Function WMIFolderExists(ByRef objRemoteWMIServer, ByVal strFolderPath, ByVal intFlag)
	'*****************************************************************************************************************************************
	'*  Purpose:				Checks to see if a folder exists (using Win32_Directory WMI call).
	'*  Arguments supplied:		Look up
	'*  Return Value:			-1 to indicate folder exists; 0 to indicate folder doesn't exist
	'*  Called by:				Main()
	'*  Calls:					None
	'*  Requirements:			ADO Constants
	'*****************************************************************************************************************************************
		Dim strFolderPathTemp, strSQLQuery, intErrNumber, strErrDescription, colFolders
		
		WMIFolderExists = False
		strFolderPathTemp = Replace(strFolderPath, "\", "\\")
		strFolderPathTemp = Replace(strFolderPathTemp, "'", "\'")
		strSQLQuery = "SELECT Name FROM Win32_Directory WHERE Name = '" & strFolderPathTemp & "'"
		Call ExecWMI(objRemoteWMIServer, intErrNumber, strErrDescription, colFolders, strSQLQuery, intFlag, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colFolders))="SWBEMOBJECTSET")) Then
			If (colFolders.Count = 0) Then
				WMIFolderExists = False
			Else
				WMIFolderExists = True
			End If
		End If
		Set colFolders = Nothing
	
	End Function

	Public Function HexToDec(ByVal strHex)
	'*****************************************************************************************************************************************
	'*  Purpose:				Convert the Hex string to a decimal string value.
	'*  Arguments supplied:		Hex string
	'*  Return Value:			Decimal string
	'*  Called by:				ResolveADIZDNSAddress, CollectRegistryData
	'*  Calls:					None
	'*****************************************************************************************************************************************
		Dim lngResult, intIndex, strDigit, intDigit, intValue
	
		lngResult = 0
		For intIndex = Len(strHex) To 1 Step - 1
			strDigit = Mid(strHex, intIndex, 1)
			intDigit = InStr("0123456789ABCDEF", UCase(strDigit)) - 1
			If (intDigit >= 0) Then
				intValue = intDigit * (16 ^ (Len(strHex) - intIndex))
				lngResult = lngResult + intValue
			Else
				lngResult = 0
				intIndex = 0 ' stop the Loop
			End If
		Next
		HexToDec = lngResult
	
	End Function

	Public Function CreateDuplicateRecordset(ByRef rsCurrent, ByRef rsNew)
	'*****************************************************************************************************************************************
	'*  Purpose:				Create a duplicate recordset (format and data)
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				EnumAllUpdateRegistryKeys
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim intCount, strColumnName, intCreateDataType, intCreateMaxLength

		If (rsNew.Fields.Count = 0) Then
			For intCount = 0 To rsCurrent.Fields.Count - 1
				strColumnName = rsCurrent(intCount).Name
				intCreateDataType = rsCurrent(intCount).Type
				intCreateMaxLength = rsCurrent(intCount).DefinedSize
				rsNew.Fields.Append strColumnName, intCreateDataType, intCreateMaxLength
			Next
		Else
			'
			' Delete any existing records
			'
			Call g_objFunctions.DeleteAllRecordsetRows(rsNew)
		End If
		If (rsNew.State <> m_adStateOpen) Then
			rsNew.Open
		End If

		If (rsCurrent.RecordCount > 0) Then
			If (Not rsCurrent.BOF) Then
				rsCurrent.MoveFirst
			End If
			While Not rsCurrent.EOF
				rsNew.AddNew
				For intCount = 0 To rsCurrent.Fields.Count - 1
					rsNew(intCount) = rsCurrent(intCount)
				Next
				rsNew.Update
				rsCurrent.MoveNext
			Wend
		End If

	End Function
	
	Public Function IsWellKnownSID(ByVal strSID, ByRef strReadableName)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if the value passed is a well-known SID in Windows
	'* 							A security identifier (SID) is a unique value of variable length that
	'*							is used to identify a security principal or security group in Windows
	'*							operating systems.  Well-known SIDs are a group of SIDs that identify
	'*							generic users or generic groups.  Their values remain constant across
	'*							all operating systems.
	'*  Arguments supplied:		Look up
	'*  Return Value:			-1 indicates it is a well-known SID (strReadableName will contain a value)
	'*							0 indicates not a well-known SID
	'*  Called by:				None
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
	
		strReadableName = ""
		Select Case UCase(strSID)
			Case "S-1-0"
				strReadableName = "NULL AUTHORITY"
			Case "S-1-0-0"
				strReadableName = "NOBODY"
			Case "S-1-1"
				strReadableName = "WORLD AUTHORITY"
			Case "S-1-1-0"
				strReadableName = "EVERYONE"
			Case "S-1-2"
				strReadableName = "LOCAL AUTHORITY"
			Case "S-1-2-0"
				strReadableName = "LOCAL"
			Case "S-1-2-1"
				strReadableName = "CONSOLE LOGON"
			Case "S-1-3"
				strReadableName = "CREATOR AUTHORITY"
			Case "S-1-3-0"
				strReadableName = "CREATOR OWNER"
			Case "S-1-3-1"
				strReadableName = "CREATOR GROUP"
			Case "S-1-3-2"
				strReadableName = "CREATOR OWNER SERVER"
			Case "S-1-3-3"
				strReadableName = "CREATOR GROUP SERVER"
			Case "S-1-3-4"
				strReadableName = "OWNER RIGHTS"
			Case "S-1-4"
				strReadableName = "NON-UNIQUE AUTHORITY"
			Case "S-1-5"
				strReadableName = "NT AUTHORITY"
			Case "S-1-5-1"
				strReadableName = "NT AUTHORITY\DIALUP"
			Case "S-1-5-2"
				strReadableName = "NT AUTHORITY\NETWORK"
			Case "S-1-5-3"
				strReadableName = "NT AUTHORITY\BATCH"
			Case "S-1-5-4"
				strReadableName = "NT AUTHORITY\INTERACTIVE"
			Case "S-1-5-6"
				strReadableName = "NT AUTHORITY\SERVICE"
			Case "S-1-5-7"
				strReadableName = "NT AUTHORITY\ANONYMOUS LOGON"
			Case "S-1-5-8"
				strReadableName = "NT AUTHORITY\PROXY"
			Case "S-1-5-9"
				strReadableName = "NT AUTHORITY\ENTERPRISE DOMAIN CONTROLLERS"
			Case "S-1-5-10"
				strReadableName = "NT AUTHORITY\PRINCIPAL SELF"
			Case "S-1-5-11"
				strReadableName = "NT AUTHORITY\AUTHENTICATED USERS"
			Case "S-1-5-12"
				strReadableName = "NT AUTHORITY\RESTRICTED CODE"
			Case "S-1-5-13"
				strReadableName = "NT AUTHORITY\TERMINAL SERVER USERS"
			Case "S-1-5-14"
				strReadableName = "NT AUTHORITY\REMOTE INTERACTIVE LOGON"
			Case "S-1-5-15", "S-1-5-17"
				strReadableName = "NT AUTHORITY\THIS ORGANIZATION"
			Case "S-1-5-18"
				strReadableName = "NT AUTHORITY\LOCAL SYSTEM"
			Case "S-1-5-19"
				strReadableName = "NT AUTHORITY\LOCAL SERVICE"
			Case "S-1-5-20"
				strReadableName = "NT AUTHORITY\NETWORK SERVICE"
			'
			' The following are Built-in Groups
			'
			Case "S-1-5-32-544"
				strReadableName = "BUILTIN\ADMINISTRATORS"
			Case "S-1-5-32-545"
				strReadableName = "BUILTIN\USERS"
			Case "S-1-5-32-546"
				strReadableName = "BUILTIN\GUESTS"
			Case "S-1-5-32-547"
				strReadableName = "BUILTIN\POWER USERS"
			Case "S-1-5-32-548"
				strReadableName = "BUILTIN\ACCOUNT OPERATORS"
			Case "S-1-5-32-549"
				strReadableName = "BUILTIN\SERVER OPERATORS"
			Case "S-1-5-32-550"
				strReadableName = "BUILTIN\PRINT OPERATORS"
			Case "S-1-5-32-551"
				strReadableName = "BUILTIN\BACKUP OPERATORS"
			Case "S-1-5-32-552"
				strReadableName = "BUILTIN\REPLICATORS"
			Case "S-1-5-32-554"
				strReadableName = "BUILTIN\PRE-WINDOWS 2000 COMPATIBLE ACCESS"
			Case "S-1-5-32-555"
				strReadableName = "BUILTIN\REMOTE DESKTOP USERS"
			Case "S-1-5-32-556"
				strReadableName = "BUILTIN\NETWORK CONFIGURATION OPERATORS"
			Case "S-1-5-32-557"
				strReadableName = "BUILTIN\INCOMING FOREST TRUST BUILDERS"
			Case "S-1-5-32-558"
				strReadableName = "BUILTIN\PERFORMANCE MONITOR USERS"
			Case "S-1-5-32-559"
				strReadableName = "BUILTIN\PERFORMANCE LOG USERS"
			Case "S-1-5-32-560"
				strReadableName = "BUILTIN\WINDOWS AUTHORIZATION ACCESS GROUP"
			Case "S-1-5-32-561"
				strReadableName = "BUILTIN\TERMINAL SERVER LICENSE SERVERS"
			Case "S-1-5-32-562"
				strReadableName = "BUILTIN\DISTRIBUTED COM USERS"
			Case "S-1-5-32-569"
				strReadableName = "BUILTIN\CRYPTOGRAPHIC OPERATORS"
			Case "S-1-5-32-573"
				strReadableName = "BUILTIN\EVENT LOG READERS"
			Case "S-1-5-32-574"
				strReadableName = "BUILTIN\CERTIFICATE SERVICE DCOM ACCESS"
			Case "S-1-5-32-575"
				strReadableName = "BUILTIN\RDS REMOTE ACCESS SERVERS"
			Case "S-1-5-32-576"
				strReadableName = "BUILTIN\RDS ENDPOINT SERVERS"
			Case "S-1-5-32-577"
				strReadableName = "BUILTIN\RDS MANAGEMENT SERVERS"
			Case "S-1-5-32-578"
				strReadableName = "BUILTIN\HYPER-V ADMINISTRATORS"
			Case "S-1-5-32-579"
				strReadableName = "BUILTIN\ACCESS CONTROL ASSISTANCE OPERATORS"
			Case "S-1-5-32-580"
				strReadableName = "BUILTIN\REMOTE MANAGEMENT USERS"
			Case "S-1-5-64-10"
				strReadableName = "NTLM AUTHENTICATION"
			Case "S-1-5-64-14"
				strReadableName = "SCHANNEL AUTHENTICATION"
			Case "S-1-5-64-21"
				strReadableName = "DIGEST AUTHENTICATION"
			Case "S-1-5-80-0"
				strReadableName = "NT SERVICES\ALL SERVICES"
			Case "S-1-5-1000"
				strReadableName = "NT AUTHORITY\OTHER ORGANIZATION"
			Case "S-1-5-83"
				strReadableName = "NT VIRTUAL MACHINE\VIRTUAL MACHINES"
			Case "S-1-16-0"
				strReadableName = "UNTRUSTED MANDATORY LEVEL"
			Case "S-1-16-4096"
				strReadableName = "LOW MANDATORY LEVEL"
			Case "S-1-16-8192"
				strReadableName = "MEDIUM MANDATORY LEVEL"
			Case "S-1-16-8448"
				strReadableName = "MEDIUM PLUS MANDATORY LEVEL"
			Case "S-1-16-12288"
				strReadableName = "HIGH MANDATORY LEVEL"
			Case "S-1-16-16384"
				strReadableName = "SYSTEM MANDATORY LEVEL"
			Case "S-1-16-20480"
				strReadableName = "PROTECTED PROCESS MANDATORY LEVEL"
			Case "S-1-16-28672"
				strReadableName = "SECURE PROCESS MANDATORY LEVEL"
		End Select
		If (Mid(UCase(strSID), 1, 7) = "S-1-5-5") Then
			strReadableName = "NT AUTHORITY\LOGON SESSION"
		End If
		If (strReadableName <> "") Then
			'
			' Found one...return
			'
			IsWellKnownSID = -1
			Exit Function
		End If
		'
		'	SID: S-1-5-domain-500		Name: Administrator
		'	SID: S-1-5-domain-501		Name: Guest
		'	SID: S-1-5-domain-502		Name: KRBTGT
		'	SID: S-1-5-domain-512		Name: Domain Admins
		'	SID: S-1-5-domain-513		Name: Domain Users
		'	SID: S-1-5-domain-514		Name: Domain Guests
		'	SID: S-1-5-domain-515		Name: Domain Computers
		'	SID: S-1-5-domain-516		Name: Domain Controllers
		'	SID: S-1-5-domain-517		Name: Cert Publishers
		'	SID: S-1-5-root domain-518	Name: Schema Admins
		'	SID: S-1-5-root domain-519	Name: Enterprise Admins
		'	SID: S-1-5-domain-520		Name: Group Policy Creator Owners
		'	SID: S-1-5-domain-533		Name: RAS and IAS Servers
		'	SID: S-1-5-21-domain-498	Name: Enterprise Read-only Domain Controllers
		'	SID: S-1-5-21-domain-521	Name: Read-only Domain Controllers
		'	SID: S-1-5-21-domain-571	Name: Allowed RODC Password Replication Group
		'	SID: S-1-5-21-domain-572	Name: Denied RODC Password Replication Group
		'	SID: S-1-5-21-domain-522	Name: Cloneable Domain Controllers
		'
		IsWellKnownSID = 0
	
	End Function

	Public Function IsWellKnownTrustee(ByVal strTrustee, ByRef strSID)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if the value passed can be converted to a well-known SID in Windows
	'* 							A security identifier (SID) is a unique value of variable length that
	'*							is used to identify a security principal or security group in Windows
	'*							operating systems.  Well-known SIDs are a group of SIDs that identify
	'*							generic users or generic groups.  Their values remain constant across
	'*							all operating systems.
	'*  Arguments supplied:		Look up
	'*  Return Value:			-1 indicates it could be converted to a well-known SID (strSID will contain a value)
	'*							0 indicates not a well-known SID
	'*  Called by:				None
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strTempTrustee

		strTempTrustee = UCase(strTrustee)				
		If (Mid(strTempTrustee, 1, 2) = "S-") Then
			strSID = strTrustee
			IsWellKnownTrustee = -1
			Exit Function
		End If
		
		strSID = ""
		Select Case UCase(strTempTrustee)
			Case "NULL AUTHORITY"
				strSID = "S-1-0"
			Case "NOBODY"
				strSID = "S-1-0-0"
			Case "WORLD AUTHORITY"
				strSID = "S-1-1"
			Case "EVERYONE"
				strSID = "S-1-1-0"
			Case "LOCAL AUTHORITY"
				strSID = "S-1-2"
			Case "CREATOR AUTHORITY"
				strSID = "S-1-3"
			Case "CREATOR OWNER"
				strSID = "S-1-3-0"
			Case "CREATOR GROUP"
				strSID = "S-1-3-1"
			Case "CREATOR OWNER SERVER"
				strSID = "S-1-3-2"
			Case "CREATOR GROUP SERVER"
				strSID = "S-1-3-3"
			Case "NON-UNIQUE AUTHORITY"
				strSID = "S-1-4"
			Case "NT AUTHORITY"
				strSID = "S-1-5"
			Case "NT AUTHORITY\DIALUP"
				strSID = "S-1-5-1"
			Case "NT AUTHORITY\NETWORK"
				strSID = "S-1-5-2"
			Case "NT AUTHORITY\BATCH"
				strSID = "S-1-5-3"
			Case "NT AUTHORITY\INTERACTIVE"
				strSID = "S-1-5-4"
			Case "NT AUTHORITY\LOGON SESSION"
				strSID = "S-1-5-5-X-Y"
			Case "NT AUTHORITY\SERVICE"
				strSID = "S-1-5-6"
			Case "NT AUTHORITY\ANONYMOUS LOGON"
				strSID = "S-1-5-7"
			Case "NT AUTHORITY\PROXY"
				strSID = "S-1-5-8"
			Case "NT AUTHORITY\ENTERPRISE DOMAIN CONTROLLERS"
				strSID = "S-1-5-9"
			Case "NT AUTHORITY\PRINCIPAL SELF"
				strSID = "S-1-5-10"
			Case "NT AUTHORITY\AUTHENTICATED USERS"
				strSID = "S-1-5-11"
			Case "NT AUTHORITY\RESTRICTED CODE"
				strSID = "S-1-5-12"
			Case "NT AUTHORITY\TERMINAL SERVER USERS"
				strSID = "S-1-5-13"
			Case "NT AUTHORITY\THIS ORGANIZATION"
				strSID = "S-1-5-15"
			Case "NT AUTHORITY\LOCAL SYSTEM", "NT AUTHORITY\SYSTEM"
				strSID = "S-1-5-18"
			Case "NT AUTHORITY\LOCAL SERVICE"
				strSID = "S-1-5-19"
			Case "NT AUTHORITY\NETWORK SERVICE"
				strSID = "S-1-5-20"
			Case "NT AUTHORITY\OTHER ORGANIZATION"
				strSID = "S-1-5-1000"
			'
			' The following are Built-in Groups
			'
			Case "BUILTIN\ADMINISTRATORS"
				strSID = "S-1-5-32-544"
			Case "BUILTIN\USERS"
				strSID = "S-1-5-32-545"
			Case "BUILTIN\GUESTS"
				strSID = "S-1-5-32-546"
			Case "BUILTIN\POWER USERS"
				strSID = "S-1-5-32-547"
			Case "BUILTIN\ACCOUNT OPERATORS"
				strSID = "S-1-5-32-548"
			Case "BUILTIN\SERVER OPERATORS"
				strSID = "S-1-5-32-549"
			Case "BUILTIN\PRINT OPERATORS"
				strSID = "S-1-5-32-550"
			Case "BUILTIN\BACKUP OPERATORS"
				strSID = "S-1-5-32-551"
			Case "BUILTIN\REPLICATORS"
				strSID = "S-1-5-32-552"
			Case "BUILTIN\PRE-WINDOWS 2000 COMPATIBLE ACCESS"
				strSID = "S-1-5-32-554"
			Case "BUILTIN\REMOTE DESKTOP USERS"
				strSID = "S-1-5-32-555"
			Case "BUILTIN\NETWORK CONFIGURATION OPERATORS"
				strSID = "S-1-5-32-556"
			Case "BUILTIN\INCOMING FOREST TRUST BUILDERS"
				strSID = "S-1-5-32-557"
			Case "BUILTIN\PERFORMANCE MONITOR USERS"
				strSID = "S-1-5-32-558"
			Case "BUILTIN\PERFORMANCE LOG USERS"
				strSID = "S-1-5-32-559"
			Case "BUILTIN\WINDOWS AUTHORIZATION ACCESS GROUP"
				strSID = "S-1-5-32-560"
			Case "BUILTIN\TERMINAL SERVER LICENSE SERVERS"
				strSID = "S-1-5-32-561"
			Case "BUILTIN\DISTRIBUTED COM USERS"
				strSID = "S-1-5-32-562"
		End Select
		If (strSID <> "") Then
			'
			' Found one...return
			'
			IsWellKnownTrustee = -1
			Exit Function
		End If
		'
		'	SID: S-1-5-domain-500		Name: Administrator
		'	SID: S-1-5-domain-501		Name: Guest
		'	SID: S-1-5-domain-502		Name: KRBTGT
		'	SID: S-1-5-domain-512		Name: Domain Admins
		'	SID: S-1-5-domain-513		Name: Domain Users
		'	SID: S-1-5-domain-514		Name: Domain Guests
		'	SID: S-1-5-domain-515		Name: Domain Computers
		'	SID: S-1-5-domain-516		Name: Domain Controllers
		'	SID: S-1-5-domain-517		Name: Cert Publishers
		'	SID: S-1-5-root domain-518	Name: Schema Admins
		'	SID: S-1-5-root domain-519	Name: Enterprise Admins
		'	SID: S-1-5-domain-520		Name: Group Policy Creator Owners
		'	SID: S-1-5-domain-533		Name: RAS and IAS Servers
		'
		IsWellKnownTrustee = 0
	
	End Function

	Public Function OctetToHexStr(ByRef arrByteOctet)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts an octet string (byte array) to a Hex string
	'*  Arguments supplied:		Look up
	'*  Return Value:			Converted Hex string
	'*  Called by:				FormatSchedule, LoadLocalUserTable, LoadLocalGroupsTable
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim intCount

		OctetToHexStr = ""
		For intCount = 1 To LenB(arrByteOctet)
			OctetToHexStr = OctetToHexStr & Right("0" & Hex(AscB(MidB(arrByteOctet, intCount, 1))), 2)
		Next

	End Function

	Public Function HexStrToDecStr(ByRef strHexSid)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts Hex string to Decimal string (SDDL) SID
	'*  Arguments supplied:		Look up
	'*  Return Value:			Converted Hex string
	'*  Called by:				All
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************

		Dim arrByteSID, intCount, strByte, strSIDOut, lngTemp, intBase, intOffset
		'
		' Anatomy of a SID:
		'	Byte Position
		'		0 : SID Structure Revision Level (SRL)
		'		1 : Number of Subauthority/Relative Identifier
		'		2-7 : Identifier Authority Value (IAV) [48 bits]
		'		8-x : Variable number of Subauthority or Relative Identifier (RID) [32 bits]
		'
		'	Example:
		'
		'		<Domain/Machine>\Administrator
		'			Pos :      0 |  1 |  2  3  4  5  6  7 |  8  9 10 11 | 12 13 14 15 | 16 17 18 19 | 20 21 22 23 | 24 25 26 27
		'		Hex Value:    01 | 05 | 00 00 00 00 00 05 | 15 00 00 00 | 81 BD 6F CC | 45 20 0B 2F | 4C 1C 82 03 | F4 01 00 00
		'		SID:        S- 1 |    |        -5         |     -21     | -3429875073 | -789258309  |  -58858572  |    -500
		'
		Const BYTES_IN_32BITS = 4
		Const SRL_BYTE = 0
		Const IAV_START_BYTE = 2
		Const IAV_END_BYTE = 7
		Const RID_START_BYTE = 8
		Const MSB = 3 'Most significant byte
		Const LSB = 0 'Least significant byte

		strSIDOut = ""
		ReDim arrByteSID(Len(strHexSid) / 2 - 1)
		'
		' Convert hex string into integer Array
		'
		For intCount = 0 To UBound(arrByteSID)
			strByte = CInt("&H" & Mid(strHexSid, 2 * intCount + 1, 2))
			strSIDOut = strSIDOut & Hex(strByte) & " "
			arrByteSID(intCount) = strByte
		Next
'		WScript.Echo strSIDOut
		'
		' Add SRL number
		'
		HexStrToDecStr = "S-" & arrByteSID(SRL_BYTE)
		'
		' Add Identifier Authority Value
		'
		lngTemp = 0
		For intCount = IAV_START_BYTE To IAV_END_BYTE
			lngTemp = lngTemp * 256 + arrByteSID(intCount)
		Next
		HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)
		'
		' Add a variable number of 32-bit subauthority or
		' relative identifier (RID) values.
		' Bytes are in reverse significant order.
		' i.e. HEX 01 02 03 04 => HEX 04 03 02 01
		' = (((0 * 256 + 04) * 256 + 03) * 256 + 02) * 256 + 01 = DEC 67305985
		'
		For intBase = RID_START_BYTE To UBound(arrByteSID) Step BYTES_IN_32BITS
			lngTemp = 0
			For intOffset = MSB to LSB Step -1
				lngTemp = lngTemp * 256 + arrByteSID(intBase + intOffset)
			Next
			HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)
		Next

	End Function

	Public Function ConvertObjectGUIDToStringGUID(ByVal strObjectGUID)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts an octet string GUID (byte array) to a Hex string GUID
	'*  Arguments supplied:		Look up
	'*  Return Value:			Formatted GUID string
	'*  Called by:				Any
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strTemp, strGUID

		strTemp = OctetToHexStr(strObjectGUID)
		strGUID = Mid(strTemp,1,8) & "-" & Mid(strTemp,9,4) & "-" & Mid(strTemp,13,4) & "-" & Mid(strTemp,17,4) & "-" & Mid(strTemp,21,12)
		ConvertObjectGUIDToStringGUID = strGUID

	End Function

	Public Function ConvertObjectSIDToStringSID(ByVal strObjectSID)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts an octet string GUID (byte array) to a Hex string GUID
	'*  Arguments supplied:		Look up
	'*  Return Value:			Formatted GUID string
	'*  Called by:				Any
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strSidHex, strSID

		'
		' Convert SID from Octet to Hex string
		'
		strSidHex = OctetToHexStr(strObjectSID)
		'
		' Convert Hex SID to Decimal SID (printable format)
		'
		strSID = HexStrToDecStr(strSidHex)
		ConvertObjectSIDToStringSID = UCase(strSID)
		
	End Function

	Public Function IsAlpha(ByRef strChar)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determine if the passed character is alphabetic
	'*  Arguments supplied:		Look up
	'*  Return Value:			True if alphabetic; False if non-alphabetic
	'*  Called by:				ProcessPreBuiltFile
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strCharacter, intErrNumber
		
		On Error Resume Next
		strCharacter = Asc(CStr(strChar))
		intErrNumber = Err.Number
		On Error GoTo 0
		
		If (intErrNumber <> 0) Then
			IsAlpha = False
		Else
			If (((strCharacter >= 65) And (strCharacter <= 90)) Or _
				((strCharacter >= 97) And (strCharacter <= 122))) Then
				IsAlpha = True
			Else
				IsAlpha = False
			End If
		End If
	
	End Function

	Public Function DBConnectionOpen(ByRef objConnection, ByVal strConnection)
	'*****************************************************************************************************************************************
	'*  Purpose:				Opens the specified database connection
	'*  Arguments supplied:		None
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					None
	'*  Requirements:			ADO Constants
	'*****************************************************************************************************************************************
		Dim intErrNumber, strErrDescription
	
		On Error Resume Next
		objConnection.Open strConnection
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
		If (objConnection.State <> m_adStateOpen) Then
			If (intErrNumber = -2147467259) Then
				WScript.Echo strConnection & " does not exist or is incorrectly configured...program abending"
			Else
				WScript.Echo "Connection to database failed...program abending"
			End If
			WScript.Echo vbTab & "Error: " & intErrNumber & "  Description: " & strErrDescription
			WScript.Quit
		End If
	
	End Function

	Public Function DBConnectionClose(ByRef objConnection)
	'*****************************************************************************************************************************************
	'*  Purpose:				Closes the specified database connection
	'*  Arguments supplied:		None
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					None
	'*  Requirements:			ADO Constants
	'*****************************************************************************************************************************************
	
		objConnection.Close
		
	End Function

	Public Function DecToBinOctet(ByVal strDec)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts the decimal value of an octet to the binary value.
	'*  Arguments supplied:		decimal value (string)
	'*  Return Value:			binary value (string)
	'*  Called by:				ConvertIPAddrToBinary
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strResult, intValue, intExp
	
		strResult = ""
		intValue = Int(strDec)
		intExp = 128
	
		While (intExp >= 1)
			If (intValue >= intExp) Then
				intValue = intValue - intExp
				strResult = strResult & "1"
			Else
				strResult = strResult & "0"
			End If
			intExp = intExp / 2
		Wend
		DecToBinOctet = strResult
	
	End Function
	
	Public Function ConvertIPAddrToBinary(ByVal strIPAddr)
	'*****************************************************************************************************************************************
	'*  Purpose:				Convert the IP Address to a 32-bit binary representation.
	'*  Arguments supplied:		IP Address in dotted notation (xxx.xxx.xxx.xxx)
	'*  Return Value:			32-bit binary representation of the IP Address.
	'*  Called by:				IPAddrComp
	'*  Calls:					DecToBinOctet
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strLine, blnISaySo, intOctets, intPOS, blnInvalid, intLen, strOctet
		Dim strChar, intStart, intEnd, strBinIPAddr, intOctetsProcessed
	
		strLine = CStr(strIPAddr)
		'
		' Do validation
		'
		blnISaySo = True
		intOctets = 0
		intPOS = 1
		While (blnISaySo) 
			intPOS = InStr(intPOS, strLine, ".")
			If (intPOS > 1) Then
				intOctets = intOctets + 1
				intPOS = intPOS + 1
			Else
				blnISaySo = False
			End If
		Wend
		blnISaySo = True
		intPOS = 1
		blnInvalid = False
		intLen = Len(strLine)
		While ((blnISaySo) And (intPOS <= intLen)) 
			strChar = Mid(strLine, intPOS, 1)
			If ((strChar <> "0") And (strChar <> "1") And (strChar <> "2") And (strChar <> "3") And (strChar <> "4") And _
				(strChar <> "5") And (strChar <> "6") And (strChar <> "7") And (strChar <> "8") And (strChar <> "9") And _
				(strChar <> ".")) Then
				blnInvalid = True
				blnISaySo = False
			End If
			intPOS = intPOS + 1
		Wend
		'
		' Only do numeric range validation of each octet if structure is correct.
		'
		If ((intOctets = 3) And (Not blnInvalid)) Then
			blnISaySo = True
			intPOS = 1
			intLen = Len(strLine)
			intStart = 1
			intEnd = intLen
			While (blnISaySo) 
				intPOS = InStr(intPOS, strLine, ".")
				If (intPOS > 1) Then
					strOctet = Mid(strLine, intStart, (intPOS - intStart))
					If (Int(strOctet) > 255) Then
						blnInvalid = True
						blnISaySo = False
					Else
						intPOS = intPOS + 1
						intStart = intPOS
					End If
				Else
					'
					' We are processing the fourth octet now.
					'
					strOctet = Mid(strLine, intStart, (intEnd - intStart + 1))
					If (Int(strOctet) > 255) Then
						blnInvalid = True
					End If
					blnISaySo = False
				End If
			Wend 
		End If
		If ((intOctets <> 3) Or (blnInvalid)) Then
			'
			' Build an address to return (000.000.000.000 in binary) if the format isn't correct.
			'
			strBinIPAddr = strBinIPAddr & DecToBinOctet(000)
			strBinIPAddr = strBinIPAddr & DecToBinOctet(000)
			strBinIPAddr = strBinIPAddr & DecToBinOctet(000)
			strBinIPAddr = strBinIPAddr & DecToBinOctet(000)
			ConvertIPAddrToBinary = strBinIPAddr
			Exit Function
		End If
		'
		' The IP address is in dotted-notation.
		'
		strBinIPAddr = ""
		intOctetsProcessed = 0
		intLen = Len(strLine)
		intStart = 1
		intEnd = intLen
		intPOS = 1
		Do While intOctetsProcessed <= 3
			intPOS = InStr(intPOS, strLine, ".")
			strOctet = ""
			If (intPOS > 1) Then
				strOctet = Mid(strLine, intStart, (intPOS - intStart))
				strBinIPAddr = strBinIPAddr & DecToBinOctet(strOctet)
				intPOS = intPOS + 1
				intStart = intPOS
			Else
				'
				' We are processing the fourth octet now.
				'
				strOctet = Mid(strLine, intStart, (intEnd - intStart + 1))
				strBinIPAddr = strBinIPAddr & DecToBinOctet(strOctet)
			End If
			intOctetsProcessed = intOctetsProcessed + 1
		Loop
		ConvertIPAddrToBinary = strBinIPAddr
		
	End Function

	Public Function IPAddrComp(ByVal strClientIPAddr, ByVal strIPSubnetMask, ByVal intMaskBits)
	'*****************************************************************************************************************************************
	'*  Purpose:				Compares a passed IP address string to a subnet mask to determine if the address is within the subnet.
	'*  Arguments supplied:		Client IP Address and Subnet Mask in dotted notation (xxx.xxx.xxx.xxx), and the number of bits to mask.
	'*  Return Value:			0 if the address is within the subnet, <> 0 if it isn't.
	'*  Called by:				BindRootDSE
	'*  Calls:					ConvertIPAddrToBinary
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim intMaskBitsLocal, strBinClientIPAddr, strBinIPSubnetMask
		Dim strCompare1, strCompare2
		
		intMaskBitsLocal = Int(intMaskBits)
		strBinClientIPAddr = ConvertIPAddrToBinary(strClientIPAddr)
		strBinIPSubnetMask = ConvertIPAddrToBinary(strIPSubnetMask)
		strCompare1 = Left(strBinClientIPAddr, intMaskBitsLocal)
		strCompare2 = Left(strBinIPSubnetMask, intMaskBitsLocal)
		IPAddrComp = StrComp(strCompare1, strCompare2)
	
	End Function

	Public Function GetProcessingSelections(ByRef rsSelections, ByVal strPrompt, ByVal strTitle, ByVal blnBack, _
												ByVal blnMultiple, ByRef arrSelectionOffsets)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get processing selections from user
	'*  Arguments supplied:		Look up
	'*  Return Value:			Always True
	'*  Local Variables:		Look below
	'*  Called by:				Mainline
	'*  Calls:					Min
	'*  Requirements:			FileIO Constants
	'*****************************************************************************************************************************************
		Dim objFSO, objShell, strParentFolder, strFormFile, objFormFile, intCount, strEntry, objIE, intSelectedCount
	
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objShell = CreateObject("WScript.Shell")
	
		strParentFolder = GetParentFolder()
		strFormFile = strParentFolder & "ProcessingSelections.htm"
		Set objFormFile = objFSO.OpenTextFile(strFormFile, m_FOR_WRITE, m_OVERWRITE_IF_EXISTENT)
		'
		' Create the HTM file to execute
		'
		objFormFile.WriteLine("<html>")
		objFormFile.WriteLine("<head>")
		'
		' The title below must be the same name as the AppActivate statement below
		'
		objFormFile.WriteLine("<title>" & strTitle & "</title>")
		objFormFile.WriteLine("<style>.errortext {color:red}")
		objFormFile.WriteLine(".hightext {color:blue}</style>")
		objFormFile.WriteLine("</head>")
		objFormFile.WriteLine("<script language='VBScript'>")
		objFormFile.WriteLine("<!--")
		objFormFile.WriteLine("Public InputComplete : InputComplete = 0")
		objFormFile.WriteLine("Public blnUserWantsOut : blnUserWantsOut = 0")
		objFormFile.WriteLine("Public arrSelected()")
		objFormFile.WriteLine("Public intArrSelectedRows")
		objFormFile.WriteLine("")
		objFormFile.WriteLine("Sub buttonOK_OnClick")
		objFormFile.WriteLine(vbTab & "Dim i")
		objFormFile.WriteLine("")
		objFormFile.WriteLine(vbTab & "intArrSelectedRows = 0")
		objFormFile.WriteLine(vbTab & "For i = 0 To Form1.Select1.length - 1")
		objFormFile.WriteLine(vbTab & vbTab & "If (Form1.Select1.item(i).selected) Then")
		objFormFile.WriteLine(vbTab & vbTab & vbTab & "ReDim Preserve arrSelected(intArrSelectedRows)")
		objFormFile.WriteLine(vbTab & vbTab & vbTab & "arrSelected(intArrSelectedRows) = Form1.Select1.item(i).value")
		objFormFile.WriteLine(vbTab & vbTab & vbTab & "intArrSelectedRows = intArrSelectedRows + 1")
		objFormFile.WriteLine(vbTab & vbTab & "End If")
		objFormFile.WriteLine(vbTab & "Next")
		objFormFile.WriteLine(vbTab & "InputComplete = 1")
		objFormFile.WriteLine("End Sub")
		objFormFile.WriteLine("")
		objFormFile.WriteLine("")
		objFormFile.WriteLine("Sub Select1_OnDblClick")
		objFormFile.WriteLine(vbTab & "Dim i")
		objFormFile.WriteLine("")
		objFormFile.WriteLine(vbTab & "intArrSelectedRows = 0")
		objFormFile.WriteLine(vbTab & "For i = 0 To Form1.Select1.length - 1")
		objFormFile.WriteLine(vbTab & vbTab & "If (Form1.Select1.item(i).selected) Then")
		objFormFile.WriteLine(vbTab & vbTab & vbTab & "ReDim Preserve arrSelected(intArrSelectedRows)")
		objFormFile.WriteLine(vbTab & vbTab & vbTab & "arrSelected(intArrSelectedRows) = Form1.Select1.item(i).value")
		objFormFile.WriteLine(vbTab & vbTab & vbTab & "intArrSelectedRows = intArrSelectedRows + 1")
		objFormFile.WriteLine(vbTab & vbTab & "End If")
		objFormFile.WriteLine(vbTab & "Next")
		objFormFile.WriteLine(vbTab & "InputComplete = 1")
		objFormFile.WriteLine("End Sub")
		objFormFile.WriteLine("")
		objFormFile.WriteLine("")
		objFormFile.WriteLine("Sub buttonEXIT_OnClick")
	 	objFormFile.WriteLine(vbTab &  "InputComplete = 1")
		objFormFile.WriteLine(vbTab &  "blnUserWantsOut = 1")
		objFormFile.WriteLine("End Sub")
		objFormFile.WriteLine("")
		objFormFile.WriteLine("")
		If (blnBack) Then
			objFormFile.WriteLine("Sub buttonBACK_OnClick")
	 		objFormFile.WriteLine(vbTab &  "InputComplete = 1")
			objFormFile.WriteLine(vbTab & "ReDim Preserve arrSelected(intArrSelectedRows)")
			objFormFile.WriteLine(vbTab & "arrSelected(intArrSelectedRows) = -1")
			objFormFile.WriteLine(vbTab & "intArrSelectedRows = intArrSelectedRows + 1")
			objFormFile.WriteLine("End Sub")
			objFormFile.WriteLine("")
			objFormFile.WriteLine("")
		End If
		objFormFile.WriteLine("Sub window_onload")
		objFormFile.WriteLine(vbTab & "Form1.elements(0).focus")
		objFormFile.WriteLine("End Sub")
		objFormFile.WriteLine("")
		objFormFile.WriteLine("")
		objFormFile.WriteLine("-->")
		objFormFile.WriteLine("</script>")
		objFormFile.WriteLine("<form name='Form1'>" & strPrompt)
		objFormFile.WriteLine(vbTab & "<br><br>")
		'
		' Limit the number of visible entries in the listbox (without a scrollbar being added) to 10
		'
		intCount = Min(rsSelections.RecordCount, 10)
		'
		' Check to see if single or multiple selections are available
		'
		If (blnMultiple) Then
			objFormFile.WriteLine(vbTab & "<select multiple id=""Select1"" name=""Select1"" size=" & intCount & " value=""Selection"">")
		Else
			objFormFile.WriteLine(vbTab & "<select id=""Select1"" name=""Select1"" size=" & intCount & " value=""Selection"">")
		End If
		intCount = 0
		If (Not rsSelections.BOF) Then
			rsSelections.MoveFirst
		End If
		While Not rsSelections.EOF
			objFormFile.WriteLine(vbTab & "<option value=" & Chr(34) & intCount & Chr(34) & ">" & rsSelections("Selection"))
			intCount = intCount + 1
			rsSelections.MoveNext
		Wend
		objFormFile.WriteLine(vbTab & "</select>")
		objFormFile.WriteLine("<br><br>")
		If (blnBack) Then
			objFormFile.WriteLine("<input type='button' name='ButtonBACK' value='Back'>&nbsp;&nbsp;&nbsp")
		End If
		objFormFile.WriteLine("<input type='button' name='ButtonOK' value='OK'>&nbsp;&nbsp;&nbsp")
		objFormFile.WriteLine("<input type='button' name='ButtonEXIT' value='Exit'>")
		objFormFile.WriteLine("</body>")
		objFormFile.WriteLine("</form>")
		objFormFile.WriteLine("</html>")
		objFormFile.Close
		Set objFormFile = Nothing
		'
		' The form file is built...use it to prompt the user for servers to process
		'
		Set objIE = WScript.CreateObject("InternetExplorer.Application")
		objIE.Left = 50
		objIE.top = 50
		objIE.height = 400
		objIE.width = 700
		objIE.menubar = 0
		objIE.toolbar = 0
		objIE.statusbar = 0
		objIE.Resizable = 0
		objIE.navigate strFormFile
		objIE.visible = 1
		Do While (objIE.Busy)
		Loop
		'
		' The title below must be the same name as the title in the HTML code above
		'
		objShell.AppActivate Chr(34) & strTitle & Chr(34)
		Do
		Loop While (objIE.document.script.InputComplete = 0)
		'
		' Load the global variables needed for processing
		'
		If (objIE.document.script.blnUserWantsOut) Then
			WScript.Echo "You selected Exit - bye bye"
			objIE.Quit()
			Set objIE = Nothing
			WScript.Quit
		End If
	
		intSelectedCount = 0
		For intCount = 0 To objIE.Document.script.intArrSelectedRows - 1
			ReDim Preserve arrSelectionOffsets(intCount)
			arrSelectionOffsets(intCount) = objIE.Document.script.arrSelected(intCount)
			intSelectedCount = intSelectedCount + 1
		Next
		'
		' Let's get out of here
		'
		objIE.Quit()
		Set objIE = Nothing
		Set objFSO = Nothing
		Set objShell = Nothing
		GetProcessingSelections = intSelectedCount
	
	End Function

	Public Function EnsureRequiredFileExists(ByRef strFilePath, ByVal strDefaultFileName, ByVal strFileType)
	'*****************************************************************************************************************************************
	'*  Purpose:				Prompts user for the specified file location
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					GetOSVersion, GetParentFolder
	'*	Requirements:			OSVersion Constants
	'*****************************************************************************************************************************************
		Dim objFSO, strOSVersion, objDialog, objBrowse, intRetVal
	
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		strOSVersion = GetOSVersion()
		If (strOSVersion = m_OS_VERSION_XP) Then
			Set objDialog = CreateObject("UserAccounts.CommonDialog")
		Else
			Set objBrowse = New ClsBrowse
		End If
	
		intRetVal = vbNo
		If (objFSO.FileExists(strFilePath)) Then
			intRetVal = MsgBox("Use " & strFilePath & " as the " & strFileType & " file?", vbYesNo Or vbDefaultButton1, "File Settings")
		End If
		If (intRetVal = vbNo) Then
			'
			' Either the file doesn't exist or the user doesn't want to use the one that exists
			'
			If (strOSVersion = m_OS_VERSION_XP) Then
				MsgBox "Please choose the " & strDefaultFileName & " file",, "Requested File"
				objDialog.InitialDir = GetParentFolder()
				objDialog.Filter = "All files|*.*"
				'
				' Open the dialog and return the selected file name
				'
				If (Not objDialog.ShowOpen) Then
					MsgBox "You didn't select the " & strDefaultFileName & " file.  This is a required file...Too bad...So sad...bye bye.",, "File Selection Processing"
					WScript.Quit
				End If
				strFilePath = objDialog.FileName
			Else
				intRetVal = objBrowse.ChooseFile("Please choose the " & strDefaultFileName & " file")
				If (intRetVal = "") Then
					MsgBox "You didn't select the " & strDefaultFileName & " file.  This is a required file...Too bad...So sad...bye bye.",, "File Selection Processing"
					WScript.Quit
				End If
				strFilePath = intRetVal
			End If
		End If
		'
		' Cleanup
		'
		Set objFSO = Nothing
		If (strOSVersion = m_OS_VERSION_XP) Then
		 	Set objDialog = Nothing
		Else
			Set objBrowse = Nothing
		End If
	
	End Function

	Public Function GetLatestLogonTime(ByVal objMachineOrUser, ByRef dtLatestLogonTime)
	'*****************************************************************************************************************************************
	'*  Purpose:				Returns the latest logon time (for one or more copies of a given user/computer)
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					Integer8Date
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim lngLastLogon, intErrNumber, dtCurrentLogonTime, lngLastLogonTimestamp, dtCurrentLogonTimestampTime

		dtLatestLogonTime = m_DEFAULT_DATE
		On Error Resume Next
		Set lngLastLogon = objMachineOrUser.Get("lastLogon")
		intErrNumber = Err.Number
		On Error GoTo 0
		If (intErrNumber = 0) Then
			'
			' Don't adjust time to local
			'
			dtCurrentLogonTime = Integer8Date(lngLastLogon, m_INTEGER8_DATE_NOADJUST)
		Else 
			lngLastLogon = 0
			dtCurrentLogonTime = m_DEFAULT_DATE
		End If
		On Error Resume Next
		Set lngLastLogonTimestamp = objMachineOrUser.Get("lastLogonTimestamp")
		intErrNumber = Err.Number
		On Error GoTo 0
		If (intErrNumber = 0) Then
			'
			' Don't adjust time to local
			'
			dtCurrentLogonTimestampTime = Integer8Date(lngLastLogonTimestamp, m_INTEGER8_DATE_NOADJUST)
		Else
			lngLastLogonTimestamp = 0
			dtCurrentLogonTimestampTime = m_DEFAULT_DATE
		End If
		'
		' If the string that is returned is not DEFAULT_DATE and the value returned is 
		' greater than the Latest value update the Latest value.  If no update occurs then the 
		' dtLatestLogonTime will be equal to 0.
		'
		If ((CStr(dtCurrentLogonTime) <> m_DEFAULT_DATE) And _
			((CDate(dtCurrentLogonTime)) > (CDate(dtLatestLogonTime)))) Then
			dtLatestLogonTime = dtCurrentLogonTime
		End If
		If ((CStr(dtCurrentLogonTimestampTime) <> m_DEFAULT_DATE) And _
			((CDate(dtCurrentLogonTimestampTime)) > (CDate(dtLatestLogonTime)))) Then
			dtLatestLogonTime = dtCurrentLogonTimestampTime
		End If

	End Function
	
	Public Function GetLogonCount(ByVal objMachineOrUser, ByRef intTotalLogonCount)
	'*****************************************************************************************************************************************
	'*  Purpose:				Returns the latest logon time (for one or more copies of a given user/computer)
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim intLogonCount, intErrNumber

		On Error Resume Next
		intLogonCount = objMachineOrUser.Get("logonCount")
		intErrNumber = Err.Number
		On Error GoTo 0
		If (intErrNumber = 0) Then
			intTotalLogonCount = intLogonCount
		End If
		
	End Function

	Public Function LoadRS(ByRef rsToLoad, ByRef strColumnName, ByRef strValue, ByRef objLogAndTrace, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Loads the recordset field (strColumnName) with the strValue (casting and bounds checking included)
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Main() and/or Load...Table routines
	'*  Calls:					LogThis, CreatePrintableString
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strValueToDisplay, strDefinedSizeToDisplay, strName, intType, intDefinedSize, blnBadData, strType, valValue, dtTemp, intErrNumber
		Dim intLen, strToUse

		Const DEFAULT_DATE = #1/1/1970#
		Const adVarChar = 200
		Const adBoolean = 11
		Const adInteger = 3
		Const adSmallInt = 2
		Const adBigInt = 20
		Const adDouble = 5
		Const adDate = 7
		Const adDBTimeStamp = 135
		Const adVarWChar = 202
		Const adLongVarChar = 201
		Const adLongVarWChar = 203
		Const adWChar = 130
		'
		' Recordsets are self-contained (for the most part).  Each column has properties
		' that are available just by asking and are as follows:
		'
		'	Name - The column name
		'	Type - The defined datatype (see ADO Constants)
		'	DefinedSize		Maximum size (in bytes, not characters) that the column can be.
		'					Although this value is present in all recordsets it has more
		'					meaning when referring to a TEXT (adVarWChar) or MEMO (adLongVarWChar)
		'					field.  An adInteger will be 4 bytes.  An adboolean will be 2 bytes.
		'					An adLongVarWChar will be 536870910 in Access (regardless of what is
		'					specified during the creation process).
		'	ActualSize		The current size (in bytes, not characters) that the column is using
		'					in THIS recordset.  If we are looking at a newly created record or an
		'					empty table, this attribute will not be available.
		'
		If (IsNull(strValue)) Then
			strValueToDisplay = "NULL"
		ElseIf (IsArray(strValue)) Then
			strValueToDisplay = Join(strValue)
		Else
			strValueToDisplay = strValue
		End If
		If (IsNull(rsToLoad(strColumnName).DefinedSize)) Then 
			strDefinedSizeToDisplay = "NULL"
		Else
			strDefinedSizeToDisplay = rsToLoad(strColumnName).DefinedSize
		End If
		Call LogThis("ColumnName: " & strColumnName & vbTab & "Value: " & strValueToDisplay & vbTab & "Data Type: " & _
						rsToLoad(strColumnName).Type & vbTab & "DefinedSize: " & strDefinedSizeToDisplay, objLogAndTrace)
		strName = rsToLoad(strColumnName).Name
		intType = rsToLoad(strColumnName).Type
		intDefinedSize = rsToLoad(strColumnName).DefinedSize
'		intActualSize = rsToLoad(strColumnName).ActualSize
		'
		' The following datatypes are currently in use within the database (SQL or Access):
		'
		'	adSmallInt		(2)
		'	adInteger		(3)
		'	adBoolean		(11)
		'	adDBTimeStamp	(135)
		'	adVarChar		(200)
		'	adLongVarChar	(201)
		'
		blnBadData = False

		If ((IsNull(strValue)) Or (Len(strValue) = 0)) Then
			blnBadData = True
			Call LogThis("Passed value " & strValue & " is NULL, EMPTY, Space, or Blank.  Defaulting to value of 0.", objLogAndTrace)
		End If

		Select Case intType
			Case adSmallInt
				strType = "Small Integer"
				If (blnBadData) Then
					rsToLoad(strColumnName) = 0
				Else
					On Error Resume Next
					valValue = CInt(strValue)
					intErrNumber = Err.Number
					On Error GoTo 0
					If (intErrNumber = 0) Then
						rsToLoad(strColumnName) = valValue
					Else
						rsToLoad(strColumnName) = 0
						Call LogThis("ColumnName: " & strColumnName & vbTab & "Value: " & strValueToDisplay & vbTab & "Data Type: " & _
								rsToLoad(strColumnName).Type & vbTab & "DefinedSize: " & strDefinedSizeToDisplay, objLogAndTraceErrors)
						Call LogThis("Error occurred converting " & strValue & " to CInt.  Defaulting to value of 0.", objLogAndTraceErrors)
					End If
				End If
			Case adBigInt
				strType = "Big Integer"
				If (blnBadData) Then
					rsToLoad(strColumnName) = 0
				Else
					rsToLoad(strColumnName) = strValue
				End If
			Case adInteger
				strType = "Long Integer"
				If (blnBadData) Then
					rsToLoad(strColumnName) = 0
				Else
					On Error Resume Next
					valValue = Clng(strValue)
					intErrNumber = Err.Number
					On Error GoTo 0
					If (intErrNumber = 0) Then
						rsToLoad(strColumnName) = valValue
					Else
						rsToLoad(strColumnName) = 0
						Call LogThis("ColumnName: " & strColumnName & vbTab & "Value: " & strValueToDisplay & vbTab & "Data Type: " & _
								rsToLoad(strColumnName).Type & vbTab & "DefinedSize: " & strDefinedSizeToDisplay, objLogAndTraceErrors)
						Call LogThis("Error occurred converting " & strValue & " to Clng.  Defaulting to value of 0.", objLogAndTraceErrors)
					End If
				End If
			Case adDouble
				strType = "Double"
				If (blnBadData) Then
					rsToLoad(strColumnName) = 0
				Else
					On Error Resume Next
					valValue = CDbl(strValue)
					intErrNumber = Err.Number
					On Error GoTo 0
					If (intErrNumber = 0) Then
						rsToLoad(strColumnName) = valValue
					Else
						rsToLoad(strColumnName) = 0
						Call LogThis("ColumnName: " & strColumnName & vbTab & "Value: " & strValueToDisplay & vbTab & "Data Type: " & _
								rsToLoad(strColumnName).Type & vbTab & "DefinedSize: " & strDefinedSizeToDisplay, objLogAndTraceErrors)
						Call LogThis("Error occurred converting " & strValue & " to CDbl.  Defaulting to value of 0.", objLogAndTraceErrors)
					End If
				End If
			Case adBoolean
				strType = "Boolean"
				If (blnBadData) Then
					rsToLoad(strColumnName) = 0
				Else
					If (strValue = False) Then
						rsToLoad(strColumnName) = 0
					Else
						rsToLoad(strColumnName) = 1
					End If
				End If
			Case adDBTimeStamp, adDate
				strType = "DateTime"
				If (blnBadData) Then
					rsToLoad(strColumnName) = DEFAULT_DATE
				Else
					On Error Resume Next
					dtTemp = CDate(strValue)
					intErrNumber = Err.Number
					On Error GoTo 0
					If (intErrNumber = 0) Then
						If ((dtTemp = "1/1/1601") Or (dtTemp = "1/1/1801")) Then
							rsToLoad(strColumnName) = DEFAULT_DATE
						Else
							rsToLoad(strColumnName) = dtTemp
						End If
					Else
						rsToLoad(strColumnName) = DEFAULT_DATE
						Call LogThis("ColumnName: " & strColumnName & vbTab & "Value: " & strValueToDisplay & vbTab & "Data Type: " & _
								rsToLoad(strColumnName).Type & vbTab & "DefinedSize: " & strDefinedSizeToDisplay, objLogAndTraceErrors)
						Call LogThis("Error occurred converting " & strValue & " to CDate.  Defaulting to value of " & DEFAULT_DATE & ".", _
										objLogAndTraceErrors)
					End If
				End If
			Case adWChar
				strType = "Char"
				If (blnBadData) Then
					rsToLoad(strColumnName) = " "
				Else
					strToUse = strValue
					Call CreatePrintableString(strToUse)
					If (Len(strToUse) > intDefinedSize) Then
						intLen = intDefinedSize
						Call LogThis("ColumnName: " & strColumnName & vbTab & "Value: " & strValueToDisplay & vbTab & "Data Type: " & _
								rsToLoad(strColumnName).Type & vbTab & "DefinedSize: " & strDefinedSizeToDisplay, objLogAndTraceErrors)
						Call LogThis("Char length (" & Len(strToUse) & ") longer than column size.  Truncating...", _
										objLogAndTraceErrors)
					Else
						intLen = Len(strToUse)
					End If
					rsToLoad(strColumnName) = Mid(strToUse, 1, intLen)
				End If
			Case adVarChar, adVarWChar
				strType = "Text"
				If (blnBadData) Then
					rsToLoad(strColumnName) = " "
				Else
					strToUse = strValue
					Call CreatePrintableString(strToUse)
					If (Len(strToUse) > intDefinedSize) Then
						intLen = intDefinedSize
						Call LogThis("ColumnName: " & strColumnName & vbTab & "Value: " & strValueToDisplay & vbTab & "Data Type: " & _
								rsToLoad(strColumnName).Type & vbTab & "DefinedSize: " & strDefinedSizeToDisplay, objLogAndTraceErrors)
						Call LogThis("Text length (" & Len(strToUse) & ") longer than column size.  Truncating...", _
										objLogAndTraceErrors)
					Else
						intLen = Len(strToUse)
					End If
					rsToLoad(strColumnName) = Mid(strToUse, 1, intLen)
				End If
			Case adLongVarChar, adLongVarWChar
				strType = "Memo"
				If (blnBadData) Then
					rsToLoad(strColumnName) = " "
				Else
					strToUse = strValue
					Call CreatePrintableString(strToUse)
					rsToLoad(strColumnName) = strToUse
				End If
			Case Else	' forgot to map a defined field type
				Call LogThis("Aaron forgot to map the field type: " & intType & "(ColumnName: " & _
								strColumnName & "  " & "Value: " & strValueToDisplay & ")", objLogAndTraceErrors)
		End Select

	End Function

	Public Function ExecCmdGenericMBSA(ByRef strQuery, ByRef rsToReturn, ByRef objLogAndTrace)
	'*****************************************************************************************************************************************
	'*  Purpose:				Initiates .Exec or cmd calls using passed query as input
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 if successful, -1 if unsuccessful.  Also loads the objRS with data if successful.  Only returns StdOut.
	'*  Called by:				Mainline
	'*  Calls:					LogThis
	'*  Requirements:			m_objLogAndTraceECG
	'*****************************************************************************************************************************************
		Dim objShell, intErrNumber, strErrDescription, objExec, strErrorInfo, strRead
	
		Call LogThis("Executing strQuery command " & strQuery & " in ExecCmdGenericMBSA", objLogAndTrace)
		Set objShell = CreateObject("WScript.Shell")
		On Error Resume Next
		Set objExec = objShell.Exec(strQuery)
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
		strErrDescription = Replace(strErrDescription, VbCrLf, "")
		'
		' If an error occurred then stop the process as we are short on resources.
		' The controlling process will take care of restart processing.
		'
		If (intErrNumber <> 0) Then
			strErrorInfo = "Error " & intErrNumber & " (" & Hex(intErrNumber) & " - " & Trim(strErrDescription) & _
								") occurred executing Exec query ExecCmdGenericMBSA."
			Call LogThis(vbTab & strErrorInfo, objLogAndTrace)
			ExecCmdGenericMBSA = -1
			Exit Function
		End If
		'
		' Load the recordset with data from objExec
		'
		Call LogThis(vbTab & "Processing StdOut in ExecCmdGenericMBSA", objLogAndTrace)
		While Not objExec.StdOut.AtEndOfStream
			strRead = Trim(objExec.StdOut.ReadLine)
			rsToReturn.AddNew
			rsToReturn("SavedData") = strRead
			rsToReturn.Update
		Wend
	
		If (rsToReturn.RecordCount = 0) Then
			ExecCmdGenericMBSA = -1
		Else
			If (Not rsToReturn.BOF) Then
				rsToReturn.MoveFirst
			End If
			While Not rsToReturn.EOF
				Call LogThis(vbTab & rsToReturn("SavedData"), objLogAndTrace)
				rsToReturn.MoveNext
			Wend
			ExecCmdGenericMBSA = 0
		End If
		Set objExec = Nothing
		Set objShell = Nothing
	
	End Function

	Public Function ExecCmdGeneric(ByRef strQuery, ByRef rsToReturn, ByRef objLogAndTrace)
	'*****************************************************************************************************************************************
	'*  Purpose:				Initiates .Exec or cmd calls using passed query as input
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 if successful, -1 if unsuccessful.  Also loads the objRS with data if successful.
	'*  Called by:				Mainline
	'*  Calls:					LogThis
	'*  Requirements:			m_objLogAndTraceECG
	'*****************************************************************************************************************************************
		Dim objShell, intErrNumber, strErrDescription, objExec, strErrorInfo, strRead
	
		Call LogThis("Executing strQuery command " & strQuery & " in ExecCmdGeneric", objLogAndTrace)
		Set objShell = CreateObject("WScript.Shell")
		On Error Resume Next
		Set objExec = objShell.Exec(strQuery)
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
		strErrDescription = Replace(strErrDescription, VbCrLf, "")
		'
		' If an error occurred then stop the process as we are short on resources.
		' The controlling process will take care of restart processing.
		'
		If (intErrNumber <> 0) Then
			strErrorInfo = "Error " & intErrNumber & " (" & Hex(intErrNumber) & " - " & Trim(strErrDescription) & _
								") occurred executing Exec query ExecCmdGeneric."
			Call LogThis(vbTab & strErrorInfo, objLogAndTrace)
			ExecCmdGeneric = -1
			Exit Function
		End If
		If (IsObject(rsToReturn)) Then
			'
			' Load the recordset with data from objExec
			'
			Call LogThis(vbTab & "Processing StdOut in ExecCmdGeneric", objLogAndTrace)
			While Not objExec.StdOut.AtEndOfStream
				strRead = Trim(objExec.StdOut.ReadLine)
				rsToReturn.AddNew
				rsToReturn("SavedData") = strRead
				rsToReturn.Update
			Wend
			'
			' Load the error data too
			'
			Call LogThis(vbTab & "Processing StdErr in ExecCmdGeneric", objLogAndTrace)
			While Not objExec.StdErr.AtEndOfStream
				strRead = Trim(objExec.StdErr.ReadLine)
				rsToReturn.AddNew
				rsToReturn("SavedData") = strRead
				rsToReturn.Update
			Wend
		
			If (rsToReturn.RecordCount = 0) Then
				ExecCmdGeneric = -1
			Else
				If (Not rsToReturn.BOF) Then
					rsToReturn.MoveFirst
				End If
				While Not rsToReturn.EOF
					Call LogThis(vbTab & rsToReturn("SavedData"), objLogAndTrace)
					rsToReturn.MoveNext
				Wend
				ExecCmdGeneric = 0
			End If
		Else
			ExecCmdGeneric = 0
		End If
		'
		' Cleanup
		'
		Set objExec = Nothing
		Set objShell = Nothing

	End Function

	Public Function CreateServerRegistryConnection(ByVal strComputerNameOrIPAddress, ByRef objRegServer, ByRef intErrNumber, _
														ByRef strErrDescription, ByRef strError, ByVal strUserID, ByVal strPassword, _
														ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Create a connection to a computers' registry (local or remote)
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 if successful, -1 if unsuccessful.
	'*  Called by:				All
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strNameSpace, strLocale, strAuthority, intSecurityFlags, objLocator, objWMIService, strHex
	
'#region <WBEM Constants>
		'
		' Flag Constants
		'
		Const WbemFlagReturnWhenComplete = 0
		Const WbemFlagForwardOnly = 32
		Const WbemFlagReturnImmediately = 16
		Const WbemFlagConnectUseMaxWait = 128			' Hex 80
		'
		' ImpersonateLevel Constants
		'
		Const WbemImpersonationLevelAnonymous = 1		' Short name: Anonymous - Hides the credentials of the caller.
														' Calls to WMI may fail with this impersonation level.
		Const WbemImpersonationLevelIdentify = 2		' Short name: Identify - Allows objects to query the credentials of the caller
														' Calls to WMI may fail with this impersonation level.
		Const WbemImpersonationLevelImpersonate = 3		' Short name: Impersonate - Allows objects to use the credentials of the caller.
														' This is the recommended impersonation level for Scripting API for WMI calls.
		Const WbemImpersonationLevelDelegate = 4		' Short name: Delegate - Windows 2000 and later:  Allows objects to permit other
														' objects to use the credentials of the caller. This impersonation will work with
														' Scripting API for WMI calls but may constitute an unnecessary security risk. 
		'
		' AuthenticationLevel Constants
		'
		Const WbemAuthenticationLevelDefault = 0		' Short name: Default - WMI uses the default Windows Authentication setting.
		Const WbemAuthenticationLevelNone = 1			' Short name: None - Uses no authentication.
		Const WbemAuthenticationLevelConnect = 2		' Short name: Connect - Authenticates the credentials of the client only when
														' the client establishes a relationship with the server.
		Const WbemAuthenticationLevelCall = 3			' Short name: Call - Authenticates only at the beginning of each call when the
														' server receives the request.
		Const WbemAuthenticationLevelPkt = 4			' Short name: Pkt - Authenticates that all data received is from the expected client.
		Const WbemAuthenticationLevelPktIntegrity = 5	' Short name: PktIntegrity - Authenticates and verifies that none of the data
														' transferred between client and server has been modified.
		Const WbemAuthenticationLevelPktPrivacy = 6		' Short name: PktPrivacy - Authenticates all previous impersonation levels And
														' encrypts the argument value of each remote procedure call.
		'
		' Error Constants
		'
		Const WBEM_E_ACCESS_DENIED = "80041003"
		Const WBEM_E_INVALID_NAMESPACE = "8004100E"
		Const WBEM_E_OUT_OF_MEMORY = "80041006"
'#endregion

		Const OS_VERSION_2K = "5.0"

		strNameSpace = "root\default"
		strLocale = ""
		strAuthority = ""
		
		If (GetOSVersion() > OS_VERSION_2K) Then
			intSecurityFlags = CLng(WbemFlagConnectUseMaxWait)
		Else
			intSecurityFlags = Null
		End If
		intErrNumber = 0
		strErrDescription = ""
		'
		' Start the WMI queries
		'
		Set objLocator = CreateObject("WbemScripting.SwbemLocator")
		objLocator.Security_.AuthenticationLevel = WbemAuthenticationLevelPktPrivacy
		objLocator.Security_.Privileges.AddAsString "SeSecurityPrivilege"
		objLocator.Security_.ImpersonationLevel = WbemImpersonationLevelImpersonate
		On Error Resume Next
		Set objWMIService = objLocator.ConnectServer(strComputerNameOrIPAddress, _
														strNameSpace, _
														strUserID, _
														strPassword, _
														strLocale, _
														strAuthority, _
														intSecurityFlags)
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
		If (intErrNumber = 0) Then
			'
			' Connect to the StdRegProv
			'
			On Error Resume Next
			Set objRegServer = objWMIService.Get("StdRegProv")
			intErrNumber = Err.Number
			strErrDescription = Err.Description
			On Error GoTo 0
			If (intErrNumber = 0) Then
				CreateServerRegistryConnection = 0
				Set objLocator = Nothing
				Exit Function
			Else
				CreateServerRegistryConnection = -1
				'
				' An error occurred - format the error text
				'
				strHex = Hex(intErrNumber)
				If ((strErrDescription = "") Or (strErrDescription = " ") Or (IsEmpty(strErrDescription))) Then
					strError = "Error " & intErrNumber & " (" & strHex & ") Occurred on StdRegProv on " & _
									strComputerNameOrIPAddress & " (Unknown Error)"
				Else
					strError = "Error " & intErrNumber & " Occurred on StdRegProv on " & _
									strComputerNameOrIPAddress & " (" & Trim(strErrDescription) & ")"
				End If
				Call LogThis(strError, objLogAndTraceErrors)
			End If
		Else
			CreateServerRegistryConnection = -1
			'
			' An error occurred - format the error text
			'
			strHex = Hex(intErrNumber)
			If ((strErrDescription = "") Or (strErrDescription = " ") Or (IsEmpty(strErrDescription))) Then
				strError = "Error " & intErrNumber & " (" & strHex & ") Occurred on ConnectServer on " & _
								strComputerNameOrIPAddress & " (Unknown Error)"
			Else
				strError = "Error " & intErrNumber & " Occurred on ConnectServer on " & _
								strComputerNameOrIPAddress & " (" & Trim(strErrDescription) & ")"
			End If
			Call LogThis(strError, objLogAndTraceErrors)
		End If
		CreateServerRegistryConnection = -1
		Set objLocator = Nothing
		
	End Function

	Public Function CreateServerConnection(ByVal strComputerNameOrIPAddress, ByRef objWMIService, ByRef intErrNumber, _
												ByRef strErrDescription, ByRef strError, ByVal strNameSpace, ByVal strUserID, _
												ByVal strPassword, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Create a connection to a computer (local or remote)
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 if successful, -1 if unsuccessful.
	'*  Called by:				All
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strLocale, strAuthority, intSecurityFlags, objLocator, strHex
		
'#region <WBEM Constants>
		'
		' Flag Constants
		'
		Const WbemFlagReturnWhenComplete = 0
		Const WbemFlagForwardOnly = 32
		Const WbemFlagReturnImmediately = 16
		Const WbemFlagConnectUseMaxWait = 128			' Hex 80
		'
		' ImpersonateLevel Constants
		'
		Const WbemImpersonationLevelAnonymous = 1		' Short name: Anonymous - Hides the credentials of the caller.
														' Calls to WMI may fail with this impersonation level.
		Const WbemImpersonationLevelIdentify = 2		' Short name: Identify - Allows objects to query the credentials of the caller
														' Calls to WMI may fail with this impersonation level.
		Const WbemImpersonationLevelImpersonate = 3		' Short name: Impersonate - Allows objects to use the credentials of the caller.
														' This is the recommended impersonation level for Scripting API for WMI calls.
		Const WbemImpersonationLevelDelegate = 4		' Short name: Delegate - Windows 2000 and later:  Allows objects to permit other
														' objects to use the credentials of the caller. This impersonation will work with
														' Scripting API for WMI calls but may constitute an unnecessary security risk. 
		'
		' AuthenticationLevel Constants
		'
		Const WbemAuthenticationLevelDefault = 0		' Short name: Default - WMI uses the default Windows Authentication setting.
		Const WbemAuthenticationLevelNone = 1			' Short name: None - Uses no authentication.
		Const WbemAuthenticationLevelConnect = 2		' Short name: Connect - Authenticates the credentials of the client only when
														' the client establishes a relationship with the server.
		Const WbemAuthenticationLevelCall = 3			' Short name: Call - Authenticates only at the beginning of each call when the
														' server receives the request.
		Const WbemAuthenticationLevelPkt = 4			' Short name: Pkt - Authenticates that all data received is from the expected client.
		Const WbemAuthenticationLevelPktIntegrity = 5	' Short name: PktIntegrity - Authenticates and verifies that none of the data
														' transferred between client and server has been modified.
		Const WbemAuthenticationLevelPktPrivacy = 6		' Short name: PktPrivacy - Authenticates all previous impersonation levels And
														' encrypts the argument value of each remote procedure call.
		'
		' Error Constants
		'
		Const WBEM_E_ACCESS_DENIED = "80041003"
		Const WBEM_E_INVALID_NAMESPACE = "8004100E"
		Const WBEM_E_OUT_OF_MEMORY = "80041006"
'#endregion

		Const OS_VERSION_2K = "5.0"

		strLocale = ""
		strAuthority = ""
		If (GetOSVersion() > OS_VERSION_2K) Then
			intSecurityFlags = CLng(WbemFlagConnectUseMaxWait)
		Else
			intSecurityFlags = Null
		End If
		intErrNumber = 0
		strErrDescription = ""
		'
		' Start the WMI queries
		'
		Set objLocator = CreateObject("WbemScripting.SwbemLocator")
		objLocator.Security_.AuthenticationLevel = WbemAuthenticationLevelPktPrivacy
		objLocator.Security_.Privileges.AddAsString "SeSecurityPrivilege"
		objLocator.Security_.ImpersonationLevel = WbemImpersonationLevelImpersonate
		On Error Resume Next
		Set objWMIService = objLocator.ConnectServer(strComputerNameOrIPAddress, _
														strNameSpace, _
														strUserID, _
														strPassword, _
														strLocale, _
														strAuthority, _
														intSecurityFlags)
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
		WScript.Sleep 1000
		If (intErrNumber = 0) Then
			CreateServerConnection = 0
		Else
			CreateServerConnection = -1
			'
			' An error occurred - format the error text
			'
			strHex = Hex(intErrNumber)
			If (strHex = WBEM_E_ACCESS_DENIED) Then
				strError = "Error " & intErrNumber & " (" & strHex & ") Occurred on " & strComputerNameOrIPAddress & " (Connection made - Access Denied)"
			ElseIf (strHex = WBEM_E_INVALID_NAMESPACE) Then
				'
				' See text at beginning of this function for instructions
				' for repairing the WMI namespace on a given computer.
				'
				strError = "Error " & intErrNumber & " (" & strHex & ") Occurred on " & strComputerNameOrIPAddress & " (Invalid Namespace)"
			ElseIf (strHex = WBEM_E_OUT_OF_MEMORY) Then
				strError = "Error " & intErrNumber & " (" & strHex & ") Occurred on " & strComputerNameOrIPAddress & " (Out of Memory)"
			Else
				If ((strErrDescription = "") Or (strErrDescription = " ") Or (IsEmpty(strErrDescription))) Then
					strError = "Error " & intErrNumber & " (" & strHex & ") Occurred on " & strComputerNameOrIPAddress & " (Unknown Error)"
				Else
					strError = "Error " & intErrNumber & " Occurred on " & strComputerNameOrIPAddress & " (" & Trim(strErrDescription) & ")"
				End If
			End If
			Call LogThis(strError, objLogAndTraceErrors)
		End If
		Set objLocator = Nothing
		
	End Function

	Public Function StripNull(ByVal strToStrip)
	'*****************************************************************************************************************************************
	'*  Purpose:				Replaces Null character in string with blank
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				CreateRegistryChecksEntity, CreateFileChecksEntity
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strWork
	
		strWork = strToStrip
		If (InStr(strWork, Null)) Then
			strWork = Replace(strWork, Null, "")
		End If
		If (IsNull(strWork)) Then
			strWork = ""
		End If
		StripNull = strWork
	
	End Function

	Public Function CreateNetworkShare(ByRef objWMIService, ByVal strSharePath, ByVal strShareName, ByVal intShareType, _
											ByVal intMaxConnections, ByVal strErrDescription)
	'*****************************************************************************************************************************************
	'*  Purpose:				Creates a Network Share on a Remote Computer
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					ExecWMI
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		'
		' Share Types:
		'
		'	0 (0x0) Disk Drive
		'	1 (0x1) Print Queue
		'	2 (0x2) Device
		'	3 (0x3) IPC
		'	2147483648 (0x80000000) Disk Drive Admin
		'	2147483649 (0x80000001) Print Queue Admin
		'	2147483650 (0x80000002) Device Admin
		'	2147483651 (0x80000003) IPC Admin
		'
		Dim objNewShare, intRetVal
	
		Set objNewShare = objWMIService.Get("Win32_Share")
		intRetVal = objNewShare.Create(strSharePath, strShareName, intShareType, intMaxConnections, strErrDescription)
		CreateNetworkShare = intRetVal
		'
		' Cleanup
		'
		Set objNewShare = Nothing
		
	End Function
	
	Public Function EnumerateNetworkShares(ByRef objWMIService, ByVal intFlag)
	'*****************************************************************************************************************************************
	'*  Purpose:				Displays Shares on a Remote Computer
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					ExecWMI
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strSQLQuery, intErrNumber, strErrDescription, colWMI, objWMI, strDriveType
		
		strSQLQuery = "SELECT AllowMaximum,Caption,MaximumAllowed,Name,Path,Type FROM Win32_Share"
		Call ExecWMI(objWMIService, intErrNumber, strErrDescription, colWMI, strSQLQuery, intFlag, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
			For Each objWMI In colWMI
				WScript.Echo "AllowMaximum: " & objWMI.AllowMaximum
				WScript.Echo "Caption: " & objWMI.Caption
				WScript.Echo "MaximumAllowed: " & objWMI.MaximumAllowed
				WScript.Echo "Name: " & objWMI.Name
				WScript.Echo "Path: " & objWMI.Path
				WScript.Echo "Type: " & objWMI.Type
			Next
		End If
	
	End Function
	
	Public Function EnumerateMappedDrives(ByRef objWMIService, ByVal intFlag)
	'*****************************************************************************************************************************************
	'*  Purpose:				Displays mapped drives on a Remote Computer
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					ExecWMI
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strSQLQuery, intErrNumber, strErrDescription, colWMI, objWMI, strDriveType
		
		strSQLQuery = "SELECT Name,Description,DriveType FROM Win32_LogicalDisk"
		Call ExecWMI(objWMIService, intErrNumber, strErrDescription, colWMI, strSQLQuery, intFlag, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
			For Each objWMI In colWMI
				WScript.Echo "Name: " & objWMI.Name
				WScript.Echo "Description: " & objWMI.Description
				WScript.Echo "DriveType: " & objWMI.DriveType
				Select Case objWMI.DriveType
					Case 2
						strDriveType = "Removable Disk"
					Case 3
						strDriveType = "Local Disk"
					Case 4
						strDriveType = "Network Drive"
					Case 5
						strDriveType = "CD"
					Case Else
						strDriveType = "Unknown"
				End Select
				WScript.Echo "DriveTypeText: " & strDriveType
				WScript.Echo ""
			Next
		End If
	
	End Function

	Public Function DoesNetworkShareExist(ByRef objWMIService, ByVal strShareName, ByVal intFlag)
	'*****************************************************************************************************************************************
	'*  Purpose:				Does the specified Share exist?
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					ExecWMI
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strSQLQuery, intErrNumber, strErrDescription, colWMI, objWMI, strDriveType
		
		strSQLQuery = "SELECT AllowMaximum,Caption,MaximumAllowed,Name,Path,Type FROM Win32_Share"
		Call ExecWMI(objWMIService, intErrNumber, strErrDescription, colWMI, strSQLQuery, intFlag, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
			For Each objWMI In colWMI
				WScript.Echo "AllowMaximum: " & objWMI.AllowMaximum
				WScript.Echo "Caption: " & objWMI.Caption
				WScript.Echo "MaximumAllowed: " & objWMI.MaximumAllowed
				WScript.Echo "Name: " & objWMI.Name
				WScript.Echo "Path: " & objWMI.Path
				WScript.Echo "Type: " & objWMI.Type
			Next
		End If
	
	End Function
	
' '====================
' 'ShareSetup.vbs
' 'Author: Jonathan Warnken - jon.warnken@gmail.com
' 'Credits: parts of various other posted scripts used
' 'Requirements: Admin Rights
' '====================
' Option Explicit 
' Const FILE_SHARE = 0
' Const MAXIMUM_CONNECTIONS = 25
' Dim strComputer
' Dim objWMIService
' Dim objNewShare
' 
' strComputer = "."
' Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
' Set objNewShare = objWMIService.Get("Win32_Share")
' Call sharesec ("C:\Robot", "Robot", "SCOT SHARE", "POSTerminalUsers_GG")

' Sub sharesec(Fname,shr,info,account)
' Dim FSO
' Dim Services
' Dim SecDescClass
' Dim SecDesc
' Dim Trustee
' Dim ACE
' Dim Share
' Dim InParam
' Dim Network
' Dim FolderName
' Dim AdminServer
' Dim ShareName
' 
' FolderName = Fname
' AdminServer = "\\" & strComputer
' ShareName = shr
' 
' Set Services = GetObject("WINMGMTS:{impersonationLevel=impersonate,(Security)}!" & AdminServer & "\ROOT\CIMV2")
' Set SecDescClass = Services.Get("Win32_SecurityDescriptor")
' Set SecDesc = SecDescClass.SpawnInstance_()
' 
' 'Set Trustee = Services.Get("Win32_Trustee").SpawnInstance_
' 'Trustee.Domain = Null
' 'Trustee.Name = "EVERYONE"
' 'Trustee.Properties_.Item("SID") = Array(1, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0)
' 
' Set Trustee = SetGroupTrustee("Retail", account) 'Use SetGroupTrustee for groups and SetAccountTrustee for users
' Set ACE = Services.Get("Win32_Ace").SpawnInstance_
' ACE.Properties_.Item("AccessMask") = 2032127 '2032127 = "Full"; 1245631 = "Change"; 1179817 = "Read"
' ACE.Properties_.Item("AceFlags") = 3
' ACE.Properties_.Item("AceType") = 0
' ACE.Properties_.Item("Trustee") = Trustee
' SecDesc.Properties_.Item("DACL") = Array(ACE)
' Set Share = Services.Get("Win32_Share")
' Set InParam = Share.Methods_("Create").InParameters.SpawnInstance_()
' InParam.Properties_.Item("Access") = SecDesc
' InParam.Properties_.Item("Description") = "Public Share"
' InParam.Properties_.Item("Name") = ShareName
' InParam.Properties_.Item("Path") = FolderName
' InParam.Properties_.Item("Type") = 0
' Share.ExecMethod_ "Create", InParam
' End Sub 
' 
' Function SetAccountTrustee(strDomain, strName) 
' Dim objTrustee
' Dim account
' Dim accountSID
' set objTrustee = getObject("Winmgmts:{impersonationlevel=impersonate}!root/cimv2:Win32_Trustee").Spawninstance_ 
' set account = getObject("Winmgmts:{impersonationlevel=impersonate}!root/cimv2:Win32_Account.Name='" & strName & "',Domain='" & strDomain &"'") 
' set accountSID = getObject("Winmgmts:{impersonationlevel=impersonate}!root/cimv2:Win32_SID.SID='" & account.SID &"'") 
' objTrustee.Domain = strDomain 
' objTrustee.Name = strName 
' objTrustee.Properties_.item("SID") = accountSID.BinaryRepresentation 
' set accountSID = nothing 
' set account = nothing 
' set SetAccountTrustee = objTrustee 
' End Function 
' 
' Function SetGroupTrustee(strDomain, strName) 
' Dim objTrustee
' Dim account
' Dim accountSID
' set objTrustee = getObject("Winmgmts:{impersonationlevel=impersonate}!root/cimv2:Win32_Trustee").Spawninstance_ 
' set account = getObject("Winmgmts:{impersonationlevel=impersonate}!root/cimv2:Win32_Group.Name='" & strName & "',Domain='" & strDomain &"'") 
' set accountSID = getObject("Winmgmts:{impersonationlevel=impersonate}!root/cimv2:Win32_SID.SID='" & account.SID &"'") 
' objTrustee.Domain = strDomain 
' objTrustee.Name = strName 
' objTrustee.Properties_.item("SID") = accountSID.BinaryRepresentation 
' set accountSID = nothing 
' set account = nothing 
' set SetGroupTrustee = objTrustee 
' End Function 
'
' 	Public Function CreateShare(ByRef objWMIService, ByVal strPath, ByVal strShareName, ByVal strErrDescription)
' 	'*****************************************************************************************************************************************
' 	'*  Purpose:				Creates a share
' 	'*  Arguments supplied:		Look up
' 	'*  Return Value:			0 to indicate success
' 	'*  Called by:				Mainline
' 	'*  Calls:					ExecWMI
' 	'*  Requirements:			None
' 	'*****************************************************************************************************************************************
' 		Dim strTempPath, colFolders, objShare, intRetVal
' 
' 		Const FILE_SHARE = 0
' 		Const MAXIMUM_CONNECTIONS = 1000
' 		'
' 		' Verify the file path exists
' 		'
' 		strTempPath = strPath
' 		If (InStr(strTempPath, "\") > 0) Then
' 			strTempPath = Replace(strTempPath, "\", "\\")
' 		End If
' 		'
' 		' Make sure we didn't add a '\' and now have 3 in a row
' 		'
' 		If (InStr(strTempPath, "\\\") > 0) Then
' 			strTempPath = Replace(strTempPath, "\\\", "\\")
' 		End If
' 		Set colFolders = objWMIService.ExecQuery("SELECT * FROM Win32_Directory Where Name = '" & strPath & "'")
' 		If (colFolders.Count = 0) Then
' 			'
' 			' The folder doesn't exist so we can't create the share.
' 			'
' 			WScript.Echo "Share path doesn't exist...Too bad...So sad...Bye Bye..."
' 			CreateShare = False
' 			Exit Function
' 		End If
' 		'
' 		' Create the Share
' 		'
' 		Set objShare = objWMIService.Get("Win32_Share")
' 		intRetVal = objShare.Create(strPath, strShareName, FILE_SHARE, MAXIMUM_CONNECTIONS, strErrDescription)
' ' 		Select Case intRetVal
' '			Case 0 : WScript.Echo "Success"
' '			Case 2 : WScript.Echo "Access denied"
' '			Case 8 : WScript.Echo "Unknown failure"
' '			Case 9 : WScript.Echo "Invalid name"
' '			Case 10 : WScript.Echo "Invalid level"
' '			Case 21 : WScript.Echo "Invalid parameter"
' '			Case 22 : WScript.Echo "Duplicate share"
' '			Case 23 : WScript.Echo "Redirected path"
' '			Case 24 : WScript.Echo "Unknown device or directory"
' '			Case 25 : WScript.Echo "Net name not found"
' '		End Select
' 		'
' 		' Does the share exist already?
' 		'
' 		If (intRetVal = 22) Then
' 			CreateShare = True
' 			Exit Function
' 		End If
' 		If (intRetVal <> 0) Then
' 			CreateShare = False
' 			Exit Function
' 		End If
' 		'
' 		' Share created successfully.  Secure it.
' 		' 
' 
' 
' 
' ' Call sharesec ("C:\Robot", "Robot", "SCOT SHARE", "POSTerminalUsers_GG")
' 
' 
' ' 
' 		Set objHold = objWMIService.Get("Win32_SecurityDescriptor")
' 		Set objSecurityDescriptor = objHold.SpawnInstance_()
' 		'
' 		' Create a trustee
' 		'
' 		Set objTrustee = objWMIService.Get("Win32_Trustee").SpawnInstance_
' 		objTrustee.Domain = Null
' 		objTrustee.Name = "EVERYONE"
' 		objTrustee.Properties_.Item("SID") = Array(1, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0)
' 
' 
' ' 
' ' Set Trustee = SetGroupTrustee("Retail", account) 'Use SetGroupTrustee for groups and SetAccountTrustee for users
' ' Set ACE = Services.Get("Win32_Ace").SpawnInstance_
' ' ACE.Properties_.Item("AccessMask") = 2032127 '2032127 = "Full"; 1245631 = "Change"; 1179817 = "Read"
' ' ACE.Properties_.Item("AceFlags") = 3
' ' ACE.Properties_.Item("AceType") = 0
' ' ACE.Properties_.Item("Trustee") = Trustee
' ' objSecurityDescriptor.Properties_.Item("DACL") = Array(ACE)
' ' Set objShare = objWMIService.Get("Win32_Share")
' ' Set InParam = objShare.Methods_("Create").InParameters.SpawnInstance_()
' ' InParam.Properties_.Item("Access") = objSecurityDescriptor
' ' InParam.Properties_.Item("Description") = "Public Share"
' ' InParam.Properties_.Item("Name") = ShareName
' ' InParam.Properties_.Item("Path") = FolderName
' ' InParam.Properties_.Item("Type") = 0
' ' objShare.ExecMethod_ "Create", InParam
' ' End Sub 
' ' 
' ' 
' ' Function SetAccountTrustee(strDomain, strName) 
' ' Dim objTrustee
' ' Dim account
' ' Dim accountSID
' ' set objTrustee = getObject("Winmgmts:{impersonationlevel=impersonate}!root/cimv2:Win32_Trustee").Spawninstance_ 
' ' set account = getObject("Winmgmts:{impersonationlevel=impersonate}!root/cimv2:Win32_Account.Name='" & strName & "',Domain='" & strDomain &"'") 
' ' set accountSID = getObject("Winmgmts:{impersonationlevel=impersonate}!root/cimv2:Win32_SID.SID='" & account.SID &"'") 
' ' objTrustee.Domain = strDomain 
' ' objTrustee.Name = strName 
' ' objTrustee.Properties_.item("SID") = accountSID.BinaryRepresentation 
' ' set accountSID = nothing 
' ' set account = nothing 
' ' set SetAccountTrustee = objTrustee 
' ' End Function 
' ' 
' ' 
' ' Function SetGroupTrustee(strDomain, strName) 
' ' Dim objTrustee
' ' Dim account
' ' Dim accountSID
' ' set objTrustee = getObject("Winmgmts:{impersonationlevel=impersonate}!root/cimv2:Win32_Trustee").Spawninstance_ 
' ' set account = getObject("Winmgmts:{impersonationlevel=impersonate}!root/cimv2:Win32_Group.Name='" & strName & "',Domain='" & strDomain &"'") 
' ' set accountSID = getObject("Winmgmts:{impersonationlevel=impersonate}!root/cimv2:Win32_SID.SID='" & account.SID &"'") 
' ' objTrustee.Domain = strDomain 
' ' objTrustee.Name = strName 
' ' objTrustee.Properties_.item("SID") = accountSID.BinaryRepresentation 
' ' set accountSID = nothing 
' ' set account = nothing 
' ' set SetGroupTrustee = objTrustee 
' ' End Function 
' '
' 
' 
' 
' 
' 
' 	End Function

	Public Function CalculateMaskBits(ByVal strMask)
	'*****************************************************************************************************************************************
	'*  Purpose:				Calculates number of mask bits based on the Mask passed
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					ExecWMI
	'*  Requirements:			None
	'*****************************************************************************************************************************************

		If (strMask = "255.255.255.255") Then
			CalculateMaskBits = 32
		ElseIf (strMask = "255.255.255.254") Then
			CalculateMaskBits = 31
		ElseIf (strMask = "255.255.255.252") Then
			CalculateMaskBits = 30
		ElseIf (strMask = "255.255.255.248") Then
			CalculateMaskBits = 29
		ElseIf (strMask = "255.255.255.240") Then
			CalculateMaskBits = 28
		ElseIf (strMask = "255.255.255.224") Then
			CalculateMaskBits = 27
		ElseIf (strMask = "255.255.255.192") Then
			CalculateMaskBits = 26
		ElseIf (strMask = "255.255.255.128") Then
			CalculateMaskBits = 25
		ElseIf (strMask = "255.255.255.0") Then
			CalculateMaskBits = 24
		ElseIf (strMask = "255.255.254.0") Then
			CalculateMaskBits = 23
		ElseIf (strMask = "255.255.252.0") Then
			CalculateMaskBits = 22
		ElseIf (strMask = "255.255.248.0") Then
			CalculateMaskBits = 21
		ElseIf (strMask = "255.255.240.0") Then
			CalculateMaskBits = 20
		ElseIf (strMask = "255.255.224.0") Then
			CalculateMaskBits = 19
		ElseIf (strMask = "255.255.192.0") Then
			CalculateMaskBits = 18
		ElseIf (strMask = "255.255.128.0") Then
			CalculateMaskBits = 17
		ElseIf (strMask = "255.255.0.0") Then
			CalculateMaskBits = 16
		End If
		CalculateMaskBits = 0
			
	End Function

	Public Function GetFunctionalLevel(ByVal intType, ByVal strDefaultNamingContext, ByVal strConfigurationNamingContext, _
											ByVal strDomainDNSName, ByRef objLogAndTrace, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines the mode of the domain/forest depending upon the intType
	'*  Arguments supplied:		Look up
	'*  Return Value:			Text description of the function level of the domain/forest
	'*  Called by:				LoadDomainPolicyTable()
	'*  Calls:					None
	'*	Requirements:			Functional Level Constants
	'*****************************************************************************************************************************************
	' ---------------------------------------------------------------
	' From the book "Active Directory Cookbook" by Robbie Allen
	' Publisher: O'Reilly and Associates
	' ISBN: 0-596-00466-4
	' Book web site: http://rallenhome.com/books/adcookbook/code.html
	'
	' http://support.microsoft.com/kb/322692
	' ---------------------------------------------------------------
		Dim objDomain, intNTMixedDomain, intMSDSBehaviorVersion, intErrNumber, strErrDescription
		
		'
		'	Windows Server 2003 Domain and Forest Functional Levels
		'
		'	Domain Functional Levels 			Forest Functional Levels 
		'	0 — Windows 2000 mixed				0 — Windows 2000
		'	0 — Windows 2000 native				0 — Windows 2000
		'	1 — Windows Server 2003 interim		1 — Windows Server 2003 interim
		'	2 — Windows Server 2003				2 — Windows Server 2003
		' 
		'	FUNCTIONAL_LEVEL_DOMAIN	= 0			FUNCTIONAL_LEVEL_FOREST	= 1
		'
		If (intType = 0) Then
			Call LogThis("Processing GetFunctionalLevel (domain) started", objLogAndTrace)
			Set objDomain = GetObject("LDAP://" & strDefaultNamingContext)
			objDomain.GetInfo
			On Error Resume Next
			intNTMixedDomain = objDomain.Get("nTMixedDomain")
			intErrNumber = Err.Number
			strErrDescription = Err.Description
			On Error GoTo 0
			
			If (intErrNumber <> 0) Then
				Call LogThis("ntMixedDomain AD attribute wasn't available - domain mode cannot be determined for domain " & _
									strDomainDNSName & ".  Processing GetFunctionalLevel failed.  Error: " & intErrNumber & _
									"  Description: " & strErrDescription, objLogAndTraceErrors)
				'
				' This variable should ALWAYS be available on all OS versions.  If
				' I get here then there is a problem and we won't really be able
				' to determine the domain mode.
				'
				intNTMixedDomain = 0
			End If
			On Error Resume Next
			intMSDSBehaviorVersion = objDomain.Get("msDS-Behavior-Version")
			intErrNumber = Err.Number
			strErrDescription = Err.Description
			On Error GoTo 0
			If (intErrNumber <> 0) Then
				'
				' If we are running in a Win2K environment then this variable will not be found
				' in the schema or as a 'Domain' variable within ADSI.  Set to 0 (Win2000) as a default.
				'
				intMSDSBehaviorVersion = 0
			End If
			'
			' Determine the Domain level
			'	
			'	Domain functional level:
			'	Windows 2000 mixed (the default in Windows Server 2003)
			'	Windows 2000 native
			'	Windows Server 2003 interim
			'	Windows Server 2003
			'	Windows Server 2008
			'	Windows Server 2008 R2
			'
			If ((intNTMixedDomain = 1) And (intMSDSBehaviorVersion = 0)) Then
				GetFunctionalLevel = "Windows 2000 Mixed"
			ElseIf ((intNTMixedDomain = 0) And (intMSDSBehaviorVersion = 0)) Then
				GetFunctionalLevel = "Windows 2000 Native"
			ElseIf ((intNTMixedDomain = 0) And (intMSDSBehaviorVersion = 1)) Then
				GetFunctionalLevel = "Windows Server 2003 Interim"
			ElseIf ((intNTMixedDomain = 0) And (intMSDSBehaviorVersion = 2)) Then
				GetFunctionalLevel = "Windows Server 2003"
			ElseIf ((intNTMixedDomain = 0) And (intMSDSBehaviorVersion = 3)) Then
				GetFunctionalLevel = "Windows Server 2008"
			ElseIf ((intNTMixedDomain = 0) And (intMSDSBehaviorVersion = 4)) Then
				GetFunctionalLevel = "Windows Server 2008 R2"
			Else
				GetFunctionalLevel = "Unknown/Unavailable"
			End If
			Call LogThis("Processing GetFunctionalLevel (domain) complete", objLogAndTrace)
		Else
			Call LogThis("Processing GetFunctionalLevel (forest) started")
			Set objDomain = GetObject("LDAP://cn=partitions," & strConfigurationNamingContext, objLogAndTrace)
			intNTMixedDomain = 0
			On Error Resume Next
			intMSDSBehaviorVersion = objDomain.Get("msDS-Behavior-Version")
			intErrNumber = Err.Number
			strErrDescription = Err.Description
			On Error GoTo 0
			
			If (intErrNumber <> 0) Then
				'
				' If we are running in a Win2K environment then this variable will not be found
				' in the schema or as a 'Domain' variable within ADSI.  Set to 0 (Win2000) as a default.
				'
				intMSDSBehaviorVersion = 0
			End If
			If (intMSDSBehaviorVersion = 0) Then
				GetFunctionalLevel = "Windows 2000"
			ElseIf (intMSDSBehaviorVersion = 1) Then
				GetFunctionalLevel = "Windows Server 2003 Interim"
			ElseIf (intMSDSBehaviorVersion = 2) Then
				GetFunctionalLevel = "Windows Server 2003"
			ElseIf (intMSDSBehaviorVersion = 3) Then
				GetFunctionalLevel = "Windows Server 2008"
			ElseIf (intMSDSBehaviorVersion = 4) Then
				GetFunctionalLevel = "Windows Server 2008 R2"
			Else
				GetFunctionalLevel = "Unknown/Unavailable"
			End If
			Call LogThis("Processing GetFunctionalLevel (forest) complete", objLogAndTrace)
		End If
	
	End Function 

End Class

Function DisplayHelpMessage()
'*****************************************************************************************************************************************
'*  Purpose:				Displays the command line help
'*  Called by:				Main
'*  Comments:				-
'*****************************************************************************************************************************************
	WScript.Echo ""
	Wscript.Echo "Used to generated target file of missing machines based on an original target list compare."
	WScript.Echo "SYNTAX: CScript " & WScript.ScriptName & " [/Switch] [/Switch]..."
	WScript.Echo ""
	WScript.Echo "  Valid Switch values:"
	WScript.Echo ""
	WScript.Echo "  /D, /Directory"
	WScript.Echo "     Name of Folder/Directory that contains the output XML files from selected processing (i.e XMLOutput)"
	WScript.Echo ""
	WScript.Echo "  /F, /File"
	WScript.Echo "     Name of original target file"
	WScript.Echo ""
	WScript.Echo "  /Version, /Ver, /V"
	WScript.Echo "     Display version of script and exit"
	WScript.Echo ""
	WScript.Echo "  /?, /Help"
	WScript.Echo "     This display"
	WScript.Echo ""
	WScript.Echo ""
	WScript.Echo "Examples:"
	WScript.Echo ""
	WScript.Echo "SYNTAX: CScript " & WScript.ScriptName & " [/Switch] [/Switch]..."
	WScript.Echo "  CScript " & WScript.ScriptName & " /D:D:\Pub\XMLOutput /F:D:\Pub\TargetList.txt"
	WScript.Echo "  CScript " & WScript.ScriptName & " /?"
	WScript.Echo ""
	WScript.Echo "NOTE: If a parameter contains a space it MUST be within quotes"
	WScript.Echo ""
	WScript.Echo "NOTE: Some command line switches are not required, they can be in any order, and"
	Wscript.Echo "      case is not important.  You do not need to include CScript in the command"
	WScript.Echo "      line if you have set CScript as the default scripting engine.  To set"
	WScript.Echo "      CScript as your default type the following at the command prompt:"
	WScript.Echo ""
	WScript.Echo "      CScript //h:cscript"
	WScript.Echo ""

End Function

Sub ProcessStartup()
'*****************************************************************************************************************************************
'*  Purpose:				Ensures that CSCRIPT.EXE is used as default script program.
'*  Arguments supplied:		None
'*  Return Value:			None
'*  Called by:				Mainline
'*  Calls:					None
'*****************************************************************************************************************************************
	Dim strWSHVersionAtLeast, objFSO, strWSHVersion, objShell, strCommand

	'
	' Ensure the correct version of Windows Script Host (WSH) is being run.
	'
	strWSHVersionAtLeast = "5.6.0.8515"
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	strWSHVersion = objFSO.GetFileVersion( WScript.FullName )
	If ( CStr( strWSHVersion ) < CStr( strWSHVersionAtLeast )) Then
		MsgBox( "The version of Windows Scripting Host (WSH) is not correct.  Download a version newer than " & strWSHVersionAtLeast & VbCrLf & _
				" from Microsoft at http://www.microsoft.com/downloads/results.aspx?displaylang=en&freeText=windows+script+host." )
		WScript.Sleep 1000
		WScript.Quit
	End If
	'
	' Force the script to run using CScript
	'
	If ( Right( UCase( WScript.FullName ), 11 ) = "WSCRIPT.EXE" ) Then
		Set objShell = CreateObject("WScript.Shell")
		strCommand = "cmd /k cscript.exe " & WScript.ScriptFullName & " //Nologo"
		objShell.Run strCommand, 1, True
		WScript.Quit
	End If
	Set objFSO = Nothing

End Sub


'
' Mainline
'
Call ProcessStartup()
Set g_objFSO = CreateObject("Scripting.FileSystemObject")
'
' Setup global access to non-logging classes
'
Set g_objFunctions = New LibraryFunctions
g_strParentFolder = g_objFunctions.GetParentFolder()
'
' Create required recordsets
'
Set g_rsXMLFileList = CreateObject("ADOR.Recordset")
g_rsXMLFileList.Fields.Append "Machine", adVarChar, 80
g_rsXMLFileList.Open

Set g_rsProcessingList = CreateObject("ADOR.Recordset")
g_rsProcessingList.Fields.Append "Machine", adVarChar, 80
g_rsProcessingList.Open
'
' Initialize variables
'
g_strFolder = ""
g_strFile = ""
g_strArgument = ""
'
' Process startup parameters (arguments)
'
For Each g_strArgument In Wscript.Arguments.Named
	'
	' Process the command line parameters
	'
	Select Case UCase(g_strArgument)
		Case "D", "DIRECTORY" 'Select Folder where .xml files are located i.e. XMLOutput
			g_strFolder = WScript.Arguments.Named(g_strArgument)
		Case "F", "FILE" 'Load in Original Target File to compare against
			g_strFile = WScript.Arguments.Named(g_strArgument)
		Case "VERSION", "VER", "V"
 			WScript.Echo SCRIPT_VERSION_MESSAGE
 			WScript.Quit
		Case "?", "HELP"
			Call DisplayHelpMessage
			WScript.Quit
		Case Else
			'
			' Unknown argument - show help and exit
			'
			WScript.Echo "Unknown parameter " & Chr(34) & g_strArgument & Chr(34) & " passed to script."
			WScript.Echo
			Call DisplayHelpMessage
			WScript.Quit
	End Select
Next
'
' Validate parameters
'
If (g_strFolder = "") Then
	WScript.Echo "Folder path not provided...abending"
	WScript.Echo "Utilize /? switch for help"
	WScript.Quit
End If
'
' Folder passed, does it exist?
'
If (Not g_objFSO.FolderExists(g_strFolder)) Then
	WScript.Echo "Specified Folder does not exist...abending"
	WScript.Quit
End If
'
' Folder exists, are there any files to process
'
Set g_objFolder = g_objFSO.GetFolder(g_strFolder)
Set g_colFiles = g_objFolder.Files
If (g_colFiles.Count = 0) Then
	WScript.Echo "Folder exists but there aren't any files to process...abending"
	WScript.Quit
End If
'
' Was file passed?
'
If (g_strFile = "") Then
	WScript.Echo "File not provided...abending"
	WScript.Echo "Utilize /? switch for help"
	WScript.Quit
End If
'
' File passed, does it exist?
'
If (Not g_objFSO.FileExists(g_strFile)) Then
	WScript.Echo "Specified File does not exist...abending"
	WScript.Quit
End If
'
' Folder and File passed - time to process
'
WScript.Echo "Folder to be processed: " & g_strFolder
'WScript.Echo "Files found: " & g_colFiles.Count
'
' Load the recordset with file names (minus the extension)
'
For Each g_objFile In g_colFiles
	g_strFileExtension = g_objFSO.GetExtensionName(g_objFile)
	'
	' Process the .xml file
	'
	If (InStr(1, g_strFileExtension, "xml", vbTextCompare) > 0) Then
		g_rsXMLFileList.AddNew
		g_rsXMLFileList("Machine") = Trim(Left(g_objFile.Name, Len(g_objFile.Name) - Len(g_strFileExtension) - 1))
		g_rsXMLFileList.Update
	End If
Next
If (g_rsXMLFileList.RecordCount = 0) Then
	WScript.Echo "Folder exists but there aren't any XML files to process...abending"
	WScript.Quit
End If
'
' Sort returned records
'
g_rsXMLFileList.Sort = "Machine"
'
' Get the list of machines in the original processing list
'
Set g_objFile = g_objFSO.OpenTextFile(g_strFile)
While Not g_objFile.AtEndOfStream
	g_strMachine = Trim(g_objFile.ReadLine)
	If (g_strMachine <> "") Then
		g_rsProcessingList.AddNew
		g_rsProcessingList("Machine") = g_strMachine
		g_rsProcessingList.Update
	End If
Wend
'
' Sort returned records
'
g_rsProcessingList.Sort = "Machine"
'
' Get the new file ready for processing
'
g_strFileExtension = g_objFSO.GetExtensionName(g_strFile)
g_strFileName = Left(g_strFile, Len(g_strFile) - Len(g_strFileExtension) - 1)
g_strNewFile = g_strFileName & "_NewTargetList." & g_strFileExtension
Set g_objNewFile = g_objFSO.CreateTextFile(g_strNewFile)

WScript.Echo "Number of files processed: " & g_colFiles.Count
WScript.Echo "Number of Original Target(s): " & g_rsProcessingList.RecordCount

g_intNumMachinesToGo = 0
g_intNumMachinesCollected = 0

If (Not g_rsProcessingList.BOF) Then
	g_rsProcessingList.MoveFirst
	While Not g_rsProcessingList.EOF
		g_strMachine = g_rsProcessingList("Machine")
		If (Not g_rsXMLFileList.BOF) Then
			g_rsXMLFileList.MoveFirst
		End If
		g_rsXMLFileList.Filter = "Machine Like '%" & g_strMachine & "%'"
		If (g_rsXMLFileList.RecordCount > 0) Then
			g_intNumMachinesCollected = g_intNumMachinesCollected + 1
		Else
			g_intNumMachinesToGo = g_intNumMachinesToGo + 1
			g_objNewFile.WriteLine(g_strMachine)
		End If
		g_rsXMLFileList.Filter = 0
		g_rsProcessingList.MoveNext
	Wend
End If
g_objNewFile.Close
WScript.Echo "Machines Left to get: " & g_intNumMachinesToGo
WScript.Echo "Machines Collected Thus far: " & g_intNumMachinesCollected
'
' Write to audit file to track process
'
g_strBuildUpdatedTargetListAudit = g_strParentFolder & "BuildUpdatedTargetList_Audit.txt"
If (g_objFSO.FileExists(g_strBuildUpdatedTargetListAudit)) Then
	Set g_objTextFile = g_objFSO.OpenTextFile(g_strBuildUpdatedTargetListAudit, FOR_APPEND)
Else
	Set g_objTextFile = g_objFSO.CreateTextFile(g_strBuildUpdatedTargetListAudit)
	g_objTextFile.WriteLine "Date_Time" & vbTab & vbTab & "Not Scanned" & vbTab & "Scanned" & vbTab & "Targeted" & vbTab & "Percentage" 
End If 
g_objTextFile.WriteLine Now & vbTab & g_intNumMachinesToGo & vbTab & vbTab & g_intNumMachinesCollected & vbTab & g_rsProcessingList.RecordCount & vbTab & vbTab & FormatPercent(g_intNumMachinesCollected/(g_rsProcessingList.RecordCount))
g_objTextFile.Close
'
' Cleanup
'
Set g_objFSO = Nothing
Set g_objFunctions = Nothing
Set g_rsXMLFileList = Nothing
Set g_rsProcessingList = Nothing
Set g_objFolder = Nothing
Set g_colFiles = Nothing
Set g_objFile = Nothing
Set g_objNewFile = Nothing
Set g_objTextFile = Nothing

WScript.Echo "Script Complete"
WScript.Quit
