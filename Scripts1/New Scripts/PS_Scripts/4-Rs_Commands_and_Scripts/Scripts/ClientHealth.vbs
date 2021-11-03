'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: SCCMValidator.vbs
'
' AUTHOR: Aaron Mueller
' DATE  : 05/29/2013
'
' COMMENT: New script to validate and correct ClientHealth settings
'
'	Version 1.0		Date: 14 June 2013
'	Modifications:
'		Initial release
'	Version 1.0a	Date: 14 August 2013
'		Modified code in UpdateService to change strState to match the current
'			state of the service (i.e. "STOPPED" when executing the Stop service 
'			section of code).
'		Updated LibraryFunctions and RegistryProcessing Classes from GMD.
'	Version 1.0b	Date: 08 October 2013
'		Modified code in BuildPath to handle an error that occurs during creation of
'			a non-existent folder.  Timing can cause this issue when using MultiThread
'			processing.
'	Version 1.0c	Date: 17 October 2013
'		Added function FormatAndSaveXMLFile to XMLProcessing Class.
'		Modified code to call FormatAndSaveXMLFile when creating the XML output so that it is standardized.
'		Reworked code in DoRepairProcessing to ensure the certificates are successfully deleted on the remote
'			machine prior to restarting the CCMEXEC service.  The Encryption certificate doesn't always Get
'			deleted even though CertUtil says it was.  A second -delstore command appears to do the trick.
'	Version 1.0d	Date: 20 October 2013
'		Replaced function GetADInfo.  The processing was not correct if the recordset was opened successfully but
'			there were no record returned.  The return value should have been 0 to indicate the call was
'			successful even if the RecordCount = 0 (No match was found).
'		Replaced function BuildDateString.  The processing was not correct if the hour was in the 12 o'clock (AM or PM) 
'			range.  Added code to ensure that the correct value was used (12 or 00).
'	Version 1.0e	Date: 25 October 2013
'		Added code to include HBSS client health.
'	Version 1.0f	Date: 04 November 2013
'		Added GetCertificateSettings function.
'		Added DeleteCertificate function.
'		Modified PossibleProcessing and RegistryProcessing classes
'	Version 1.0g	Date: 18 November 2013
'		Modified code in GetCertificateSettings to only use Certutil.exe to get the Certificate information.
'		Modified code in DeleteCertificate to read the registry and delete the SMS Certificate.  It is much
'			quicker to do and more reliable.
'		Modified code in GetClientConfiguration to handle processing of the CatalogVersion parameter in the 
'			agent.ini file where the CatalogVersion = "0".
'	Version 1.0h	Date: 20 November 2013
'		Added code in GetClientConfiguration to use the correct location of the agent.ini file depending on OSVersion.
'	Version 1.0i	Date: 04 December 2013
'		Modified code in ProcessXMLFiles to rename columns "WUState" to "WUAState", "WUStartMode" to "WUAStartMode", and to
'			add the "CreationTimestamp" field.
'	Version 1.0j	Date: 10 December 2013
'		Replaced code that displays "Too Bad...So Sad...Bye Bye" with something more appropriate.
'		Added code to gather the LastLoggedOnUser.
'		Modified code in DoRepairProcessing to check to see if the Service exists ("Not Installed") prior to attempting to
'			start/stop/modify the Service.
'		Modified code in GetClientConfiguration to get the SiteList from the SiteList file in 
'			c:\ProgramData\McAfee\Common Framework instead of the registry.  The information in the SiteList file
'			will be in the same order on all systems depending on the Framework package that was used for installation.  
'			Each each entry contains an "Order" attribute that shows the order a client will contact the server(s).
'		Added code in LoadUsersElement, LoadGroupsElement, LoadOUsElement, LoadADServersElement, and LoadADWorkstationsElement
'			to call RemoveInvalidCharacters for the DistinguishedName field.
'		Modified code in DeleteRemoteRegistryValue and DeleteRemoteRegistryKey (RegistryProcessing Class).
'		Added Registry processing for Component Based Servicing and Auto Update.  They indicate that a reboot is required on
'			the local machine.
'	Version 1.0k	Date: 16 December 2013
'		Modified code to work with certificates in registry if at all possible.
'		Added code in GetClientConfiguration to only check on Service information for GPClient if Vista or above.
'	Version 1.0l	Date: 02 January 2014
'		Modified code in function GetClientConfiguration to call GetCertificateSettingsSMS to get the SMS certificates.
'		Modified code in CorrectCertificate to accept the registry key or certificate number of the key and process accordingly.
'			This will ensure that multiple reads aren't required since we will know which certificate is bad and how we found
'			it during the original lookup operation.
'		Modified code in function DoRepairProcessing to call CorrectCertificate with updated parameter list.
'		Modified code to separate SCCM and HBSS XML processing into 2 separate .tsv files for distribution to a specific groups.
'	Version 1.0m	Date: 05 January 2014
'		Created recordset g_rsClientConfiguration to hold all client settings (vs passing 30+ parameters in a function call).
'	Version 1.0n	Date: 09 January 2014
'		Modified code in function ParseRegistryCertificate to correctly parse any certificate from the registry.  The format was
'			slightly different in the lab than those on the AFNET.  I have tested against my local computer and one in the lab
'			and it appears that my new parsing routine is working correctly.  Only time will tell...
'		Modified code in all of the "Certificate Functions" to use the global recordset g_rsCertificates to store any
'			certificates found during processing and then loads them all at once upon completion of collection.
'	Version 1.0p	Date: 14 January 2014
'		Modified code in function ProcessXMLFiles and ProcessXMLFile to separate processing for SCCM and HBSS to only
'			create and populate the output files if the parameter to do so is passed.
'	Version 1.0q	Date: 04 February 2014
'		Modified code in function ProcessXMLFile to move LastBootupTime and add DaysSinceLastBootup in the output structure 
'			for the report.  Removed call to Close for the SCCM Output File (throwing an error).
'	Version 1.0r	Date: 07 February 2014
'		Modified code in function GetClientConfiguration to check the return value from AVDatDate processing.  Some registry values
'			are Null, Empty, Blank or Space.  If that is the case then loading with DEFAULT_DATE.
'		Modified code in DoRepairProcessing to exit out of function if there was an error returned in the call to 
'			GetClientConfiguration.  There is already error handling in that function so 2 different errors were being produced.
'	Version 1.0s	Date: 11 February 2014
'		Added Constant vbDateMicrosoftDateTime to VerifyAndLoad Constants.
'		Modified code in function VerifyAndLoad to better handle processing of Strings that contain DateTime information.
'		Collecting SMS Unique Identifier, Previous SMSUID, and Last SMSUID Change Date from the SMSCFG.ini file.  Added the newly
'			collected fields to the .XML output file and also to the spreadsheet processing section.
'		Updated code in GetClientConfiguration to collect the following settings from SMSConfig: strMachineSID, strSMBIOSSerialNumber, 
'			strHardwareIdentifier, strHardwareIdentifier2, and strLastVersion.
'	Version 1.0t	Date: 12 February 2014
'		Modified code in function GetADInfo to update the objCommand Property "Chase Referrals" to always chase them.  The 
'			default is never chase referrals.  We want the machine where the request is issued to get us the answer, even
'			if it has to ask someone else.
'	Version 1.0u	Date: 13 February 2014
'		Modified code in function GetClientConfiguration to correctly pull the agent.ini and sitelist.ini file information from
'			C:\ProgramData\McAfee\Common Framework on Windows Vista, 7, and 2008.
'	Version 1.0v	Date: 18 February 2014
'		Modified code in function UpdateService to check the return value of intError=0, the TypeName of colWMI is SWEBMOBJECTSET, 
'			see if colWMI is Empty, and ensure that colWMI has a count > 0.
'	Version 1.0w	Date: 04 March 2014
'		Modified code in Mainline section to check g_blnRepairSiteCode = False (code was there but didn't check for = False).
'			Reported by MSgt Ziwisky at Luke AFB.
'	Version 1.0x	Date: 10 March 2014
'		Added additional items to be gathered: Current ManagementPoint, ADSiteName, Services CCMSetup, RPCSS, and LanManServer, 
'			WUAgent Version, and DCOM Enabled.
'	Version 1.0y	Date: 04 April 2014
'		Modified g_rsClientConfiguration fields GPClientState, CurrentMP, and WUAVersion from 10, 50, 20 to 20, 80, 75 respectively.
'			These correspond to database/client collection field size issued.
'	Version 1.0z	Date: 14 April 2014
'		Modified code in GetClientConfiguration to handle errors when querying information in the root\ccm namespace.  It appears 
'			there is something wrong with WMI (0x80041010 - Invalid Class).  The queries that fail are ("SELECT * FROM SMS_Authority")
'			and ("SELECT * FROM CCM_ADSiteInfo").
'	Version 1.1		Date: 16 April 2014
'		Modified code to handle processing for "SMSTSMGRState" and "SMSTSMGRStartMode" (SMS Task Sequence Agent) in Scan, Repair, and 
'			XML Processing.
'	Version 1.1a	Date: 24 April 2014
'		Modified code to split SCCM XML report information into 2 different types: Normal and Verbose.  The Verbose report contains 
'			ALL of the collected SCCM information.  The Normal report only contains information that is pertinent to ClientHealth
'			specific settings.
'		Added Service RemoteRegistry to processing.
'		ReGrouped the report fields to put WMI and WMIRegistry boolean values next to the WinMgmts Service, and RemoteRegistry boolean
'			value next to RemoteRegistry Service.
'	Version 1.1b	Date: 30 April 2014
'		Modified code to add SCCMHealthy and HBSSHealthy flags to ClientHealth table.
'		Modified code to ensure the new flags were included in the XML output file and the Excel Spreadsheet output.
'	Version 1.1c	Date: 12 May 2014
'		Modified code to collect EncryptionNotBefore, EncryptionNotAfter, SigningNotBefore, and SigningNotAfter and include in XML 
'			files and reporting in Excel.  The NotBefore dates should be before todays date, and the NotAfter dates should be after
'			todays date (meaning that the Certificate is currently valid).  If the Certificate isn't valid it should be deleted and
'			recreated (just as if the Subject was incorrect).
'	Version 1.1d	Date: 21 May 2014
'		Modified code in function BuildDateString to handle military date/time setting (24-hour clock).
'	Version 1.1e	Date: 4 June 2014
'		Modified code in function GetClientConfiguration to use library function ExecWMI for SMS_Authority and CCM_ADSiteInfo calls.
'		Ensured that wbemFlagReturnWhenComplete was passed for the iFlags parameter (SWbemServices.ExecQuery) to ensure call completes 
'			prior to return to calling function.  The program was failing due to the early return (intError wasn't being set but an 
'			error was actually occurring - due to returning prior to the call completing).
'	Version 1.1f	Date: 5 June 2014
'		Added LoggingAndTracing logic.
'		Modified all library calls to make use of global LoggingAndTracing variables.
'		Modified code in GetClientNetworkInfo (ClientResolution class) to check the HostName returned from CheckClientDNS and
'			CheckClientPING to ensure that a value is present.  This fixes a bug where the PING response is "Pinging x.x.x.x with 32 
'			bytes of data:" so the HostName is "".  This causes the calling function to think that the HostName was resolved because
'			the logic was setting the HostName to "" and HostNameResolved to True.  The blank HostName was used in a WMI connection
'			attempt and was successful (to the local machine since there wasn't a valid value).
'	Version 1.1g	Date: 24 June 2014
'		Modified code in GetClientNetworkInfo to call CheckClientDNS_NSLookup instead of CheckClientDNS (no longer exists).
'		Modified code in CheckClientDNS_DNSQuery to use 0 as the value for intFlag in the call to ExecWMI.
'	Version 1.1h	Date: 15 Jul 2014
'		Modified size of Subject field in g_rsCertificates from 150 to 255 (to match length in Certificates table).
'		Modified size of EncryptionSubject field in g_rsClientConfiguration from 50 to 255 (to match the Subject field in g_rsCertificates).
'		Modified size of SigningSubject field in g_rsClientConfiguration from 50 to 255 (to match the Subject field in g_rsCertificates).
'	Version 1.1i	Date: 30 July 2014
'		Added new function GetAdministratorsGroupName in class PossibleProcessing.  It determines the actual name of the Administrators
'			group (SID=S-1-5-32-544) on the local machine (it can and has been renamed;hence the need for this fix).
'		Modified code in ValidateProcessingOpportunities to determine the name of the local administrators group by calling 
'			GetAdministratorsGroupName.
'		Modified code in IsLocalAccountProcessingPossible to use the local administrator group name returned from the call to
'			GetAdministratorsGroupName.
'	Version 1.1j	Date: 31 Jul 2014
'		Replaced class ClientResolution with newly updated one that contains better logging and fixes this issue where multiple 
'			calls were made to determine the resolution.
'	Version 1.1k	Date: 7 Aug 2014
'		Added function GetActiveIPAddress to ClientResolution Class.
'		Modified code in GetClientConfiguration and BuildClientHealthXML functions to collect and load information regarding when the
'			last SMS patch/update was received and applied.
'	Version 1.1l	Date: 12 Aug 2014
'		Modified function RequireTwoRegistryReads in Class RegistryProcessing to remove the limitation of one read to the following
'			registry location: "HKLM\Software\Network Associates\ePolicy Orchestrator\Application Plugins".  There is data in both
'			the 32-bit and 64-bit registry locations and requires 2 reads.
'	Version 1.1m	Date: 20 Aug 2014
'		Added code to functions ProcessXMLFiles and ProcessXMLFile to include determination of "SCCMClientInstallation" by interrogating the
'			CCMSetup Service State for "UnKnown" and "Not Installed", and SCCMVersion <> "" (meaning it has been installed).  This will allow
'			easier interrogation by users in the field.
'	Version 1.1n	Date: 21 Aug 2014
'		Replaced all calls to ExecQuery with call to ExecWMI Library function.  This function contains standard error handling.
'	Version 1.1.3.0		Date: 25 August 2014
'		Modified version to comply with SCOPE EDGE 4 digit versioning information from Mike West
'	Version 1.1.3.1		Date: 5 September 2014
'		Modified code in functions EnumWMIRegistryProcessing and ExactEntryRegistryProcessing (Class RegistryProcessing) to include the
'			strErrDescription in the error message.  This will assist in troubleshooting.
'		Modified code in GetActiveIPAddress (Class ClientResolution) to check the GatewayCostMetric for IsArray.  If it is an Array then 
'			this is the connection being used (Array(0) should equal zero).
'		Modified code in GetClientConfiguration to get the SMS version so the that the location of the SCCM Cache Folder could be determined.
'			The location is different for SCCM 2007 (64 and 32 bit versions) and SCCM 2012 (same for 64 and 32 bit versions).
'	Version 1.1.3.2		Date: 8 September 2014
'		Modified code in function DoRepairProcessing to Dim variable intRetVal.
'		Modified code in UpdateService to Dim variable blnStarted.
'		Modified code in CorrectCertificate to Dim variable blnProcessed.
'		Modified intError to intErrNumber and strDescription to strErrDescription.
'	Version 1.1.3.3		Date: 15 September 2014
'		Modified code in function DoRepairProcessing to Dim variable intRetVal.
'	Version 1.1.3.4		Date: 1 October 2014
'		Modified code in functions ProcessXMLFile and ProcessXMLFiles to include PendingReboot for HBSS output.
'	Version 1.1.3.5		Date: 16 October 2014
'		Added Constant vbDateWMIDateTime to VerifyAndLoad Constants and LibraryFunctions Class.
'		Modified function VerifyAndLoad (in LibraryFunctions class) to process vbDateWMIDateTime values.
'		Modified code to use vbDateWMIDateTime constant where applicable (vs vbDate).
'		Added collection of EPO Registry information to GetClientConfiguration.
'		Added EPO Registry information to XML output file in BuildClientHealthXML.
'		Added the following fields to ClientHealth table: SiteListPort, EPORegistryName, EPORegistryIP, EPORegistryPort, AgentGUID.
'		Added the 5 new fields to ProcessXMLFiles and ProcessXMLFile.
'	Version 1.1.3.6		Date: 28 October 2014
'		Added Constant vbDateUTCDateTime to VerifyAndLoad Constants and LibraryFunctions Class.
'		Modified function VerifyAndLoad (in LibraryFunctions class) to process vbDateUTCDateTime values.
'		Created new function RemoveControlCharacters in LibraryFunctions Class.  It will accept a string and reformat it removing 
'			any control characters (ox00-0x31, 0x7F, 0x80-0x9F).
'		Added code in VerifyAndLoad function to call RemoveControlCharacters for all String data types.
'		Deleted function RemoveInvalidCharacters from LibraryFunctions Class.
'		Added code in VerifyAndLoad function to handle vbDateDNSTimestamp, vbDateMicrosoftDateTime, vbDateWMIDateTime, and vbDateUTCDateTime
'			if an Empty or Null value was passed, but it was supposed to be a Date field.  This will allow DEFAULT_DATE to be loaded in the
'			return field.
'	Version 1.1.3.7		Date: 28 November 2014
'		Added function GetRemoteOSVersion to LibraryFunctions Class.
'		Added function LogIt to LibraryFunctions Class.
'	Version 1.1.3.8		Date: 12 December 2014
'		Added function LogItRecordset to LibraryFunctions Class.
'	Version 1.1.3.9		Date: 16 December 2014
'		Modified code in function GetActiveIPAddress (ClientResolution Class) to look at the IPv4Route Table entry where 
'			Network Destination and Netmask are both 0.0.0.0 (default route).  The value for the interface attribute is the 
'			IP Address of the inuse adapter.
'	Version 1.1.3.10	Date: 13 January 2015
'		Modified code in function ChooseFile (ClsBrowse Class) to attempt creation of object InternetExplorer.Application
'			ten (10) times before failure occurs.  Also added code to gracefully handle the failure.
'	Version 1.1.3.11	Date: 10 February 2015
'		Modified code in ChooseFile to call CheckWordInstallation, use Word, then quit when done.  Multiple copies of Word
'			were being left on systems after applications completed.
'	Version 1.1.3.12	Date: 10 February 2015
'		Modified function GetActiveIPAddress in Class ClientResolution to call LogThis that is within the Class.  It was calling it
'			via objLogAndTrace.LogThis and was failing if objLogAndTrace wasn't passed as a parameter.  The LogThis function is coded
'			to handle the case where a Trace object isn't passed.
'	Version 1.1.3.13	Date: 16 February 2015
'		Added ProcessingLocal to Class LibraryFunctions.
'		Added GetLocalClientNetworkInfo to Class ClientResolution.
'	Version 1.1.3.14	Date: 20 February 2015
'		Modified function CreateNewTableElement (XMLProcessing Class) to correctly build the XML element.  Somewhere during consolidation 
'			of Classes this function was broken.
'		Modified function ProcessingLocal (LibraryFunctions Class) to return the local machine name if there was no data passed at startup.
'	Version 1.1.3.15	Date: 18 March 2015
'		Modified code to use different registry location for BuildInfo: HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation\Model
'			instead of HKLM\Software\USAF\SDC\ImageRev\CurrentBuild (both locations contain the same information).
'
'==========================================================================
Option Explicit

'#region <Global Constants>

Const SCRIPT_VERSION_MESSAGE		= "Executing ClientHealth Script Version 1.1.3.15"
Const SCRIPT_VERSION				= "1.1.3.15"

Const DEFAULT_DATE = #1/1/1970#
Const wbemFlagReturnWhenComplete = 0

'#endregion

'#region <VerifyAndLoad Constants>

Const vbGUID = 100
Const vbSID = 101
Const vbScheduleOrRelay = 102
Const vbDateDNSTimestamp = 103
Const vbDateMicrosoftDateTime = 104
Const vbDateWMIDateTime = 105
Const vbDateUTCDateTime = 106

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
Const adCmdText				= 1		' Source holds command text (e.g. a SQL string)
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

'#region <OSVersion Constants>

'
' OS versions
'
Const OS_VERSION_2K			= "5.0"
Const OS_VERSION_XP			= "5.1"
Const OS_VERSION_XP_64		= "5.2"
Const OS_VERSION_2K3		= "5.2"
Const OS_VERSION_VISTA		= "6.0"
Const OS_VERSION_2K8		= "6.0"
Const OS_VERSION_2K8_R2		= "6.1"
Const OS_VERSION_WIN7		= "6.1"
Const OS_VERSION_WIN8		= "6.2"
Const OS_VERSION_2K12		= "6.2"
Const OS_VERSION_WIN8DOT1	= "6.3"
Const OS_VERSION_2K12_R2	= "6.3"

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

'#region <Registry Constants>

'
' Registry Hive constants
'
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
'
' CheckAccess Registry constants
'
Const REG_KEY_QUERY = &H0001
Const REG_KEY_SET = &H0002
Const REG_KEY_CREATE_SUB_KEY = &H0004
Const REG_KEY_ENUMERATE_SUB_KEYS = &H0008
Const REG_KEY_NOTIFY = &H0010
Const REG_KEY_CREATE_LINK = &H0020
Const REG_KEY_WOW64_32KEY = &H0200		' Indicates that an application on 64-bit Windows should operate on the 32-bit registry view
Const REG_KEY_WOW64_64KEY = &H0100		' Indicates that an application on 64-bit Windows should operate on the 64-bit registry view
Const REG_KEY_DELETE = &H00010000
Const REG_READ_CONTROL = &H00020000
Const REG_WRITE_DAC = &H00040000
Const REG_WRITE_OWNER = &H00080000
'
' Registry Data Type constants
'
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7
Const REG_QWORD = 11
'
' Registry Read Return constants
'
Const REG_KEY_DISABLED	= 0
Const REG_KEY_ENABLED	= 1
Const REG_KEY_INITIALIZED	= -1
Const REG_KEY_DOESNT_EXIST	= -2
Const REG_KEY_NOT_CONFIGURED	= -3
Const REG_KEY_NO_READ_ACCESS	= -4
Const REG_KEY_NOT_AVAILABLE_PRIOR_TO_XP	= -5
Const REG_KEY_NOT_INSTALLED	= -6

Const REG_ODBC_INI = "SOFTWARE\ODBC\ODBC.INI\"
Const REG_ODBCINST_INI = "SOFTWARE\ODBC\ODBCINST.INI\"
Const REG_ODBC_INI_DATA_SOURCES = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"

Const REG_STANDARD_TCPIP_PORTS = "SYSTEM\CurrentControlSet\Control\Print\Monitors\Standard TCP/IP Port\Ports"

Const NIC_CLASS_KEY_NAME = "SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}"
Const NIC_NETWORK_KEY_NAME = "SYSTEM\CurrentControlSet\Control\Network\{4D36E972-E325-11CE-BFC1-08002bE10318}"
Const TCPIP_ADAPTERS_KEY_NAME = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Adapters"
Const TCPIP_INTERFACES_KEY_NAME = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces"

'#endregion

'#region <Logging Constants>

Const CLIENTHEALTH_ERROR_FILE						= "ClientHealth_Error.txt"
Const CLIENTHEALTH_TRACE_FILE						= "ClientHealth_Trace.txt"
Const CLIENTHEALTH_TRACE_CLIENT_RESOLUTION_FILE		= "ClientHealth_Trace_Client_Resolution.txt"
Const CLIENTHEALTH_TRACE_POSSIBLE_PROCESSING_FILE	= "ClientHealth_Trace_Possible_Processing.txt"
Const CLIENTHEALTH_TRACE_REGISTRY_PROCESSING_FILE	= "ClientHealth_Trace_Registry_Processing.txt"
Const CLIENTHEALTH_TRACE_EXEC_CMD_GENERIC_FILE		= "ClientHealth_Trace_Exec_Cmd_Generic.txt"
Const CLIENTHEALTH_TRACE_LOADRS_FILE				= "ClientHealth_Trace_LoadRS.txt"
Const CLIENTHEALTH_TRACE_XML_FILE_AND_REGISTRY		= "ClientHealth_Trace_XML_File_And_Registry.txt"

'#endregion

'
' Global variables
'
Dim g_objFSO, g_objFunctions, g_objXMLProcessing, g_objPossibleProcessing, g_objClientResolution, g_objRegistryProcessing, g_xmlDoc
Dim g_xmlElementPassedParams, g_xmlElementConnectionStatus, g_strConnectionStatusGUID, g_strMachine, g_blnPassedMachine, g_strSourceFile
Dim g_blnPassedFile, g_blnRepair, g_blnRepairSiteCode, g_blnScan, g_blnProcessXML_SCCM, g_blnProcessXML_HBSS, g_strSiteCode, g_blnVerbose
Dim g_blnTraceAll, g_blnTraceBasic, g_blnTracePossibleProcessing, g_blnTraceRegistryProcessing, g_blnTraceClientResolution
Dim g_blnTraceLoadRS, g_blnTraceExecCmdGeneric, g_blnTraceXMLFileAndRegistry, g_strArgument, g_dtCreationTimestamp, g_objShell
Dim g_blnProcessingLocal, g_strThisComputer, g_strParentFolder, g_rsMachinesToProcess, g_rsGeneric, g_rsClientConfiguration
Dim g_rsConnectionStatus, g_rsCertificates, g_objSourceFile, g_intFlag, g_strXMLOutputPath, g_strErrorOutputPath, g_strTraceOutputPath
Dim g_strLoggingFile, g_strGUID, g_objFolder, g_colFiles, g_strRepairOutputPath
'
' Log And Trace variables
'
Dim g_blnCreateNewFile, g_strLogAndTrace, g_strLogAndTraceErrors, g_strLogAndTraceClientResolution, g_strLogAndTracePossibleProcessing
Dim g_strLogAndTraceRegistryProcessing, g_strLogAndTraceExecCmdGeneric, g_strLogAndTraceLoadRS, g_strLogAndTraceXMLFileAndRegistry
Dim g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceClientResolution, g_objLogAndTracePossibleProcessing
Dim g_objLogAndTraceRegistryProcessing, g_objLogAndTraceExecCmdGeneric, g_objLogAndTraceLoadRS, g_objLogAndTraceXMLFileAndRegistry


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

Class LoggingAndTracing
	'
	' Usage:
	'	strTraceFile = "F:\Scripts\CurrentWork\DIGS\DIGS-SQL\VBScriptIncludeCode\junk.txt"
	'	Dim LogAndTrace : Set LogAndTrace = (New LoggingAndTracing)(strTraceFile, True, True)
	'
	'	or
	'
	'	strTraceFile = ""
	'	Dim LogAndTrace : Set LogAndTrace = (New LoggingAndTracing)(strTraceFile, False, True)
	'
	Private m_FOR_WRITE, m_FOR_APPEND, m_CREATE_IF_NON_EXISTENT, m_strLogFile, m_objFSO, m_objLogFile
	Private m_blnEnabled, m_blnCreateNewFile, m_strTextLocal

	Public Default Function Init(LogFile, Enabled, CreateNewFile) 'Constructor
		If (IsObject(g_objFunctions) = False) Then
			WScript.Echo "Object g_objFunctions required for Class LoggingAndTracing.  Abending..."
			WScript.Quit
		End If
		m_FOR_WRITE	= 2
		m_FOR_APPEND = 8
		m_CREATE_IF_NON_EXISTENT = True
		m_strLogFile = LogFile
		m_blnEnabled = Enabled
		m_blnCreateNewFile = CreateNewFile
		If (m_blnEnabled) Then
			Set m_objFSO = CreateObject("Scripting.FileSystemObject")
			If (m_blnCreateNewFile) Then
				If(m_objFSO.FileExists(m_strLogFile)) Then
					m_objFSO.DeleteFile m_strLogFile, True
				End If
				'
				' Open for Write and create if non-existent
				'
				Set m_objLogFile = m_objFSO.OpenTextFile(m_strLogFile, m_FOR_WRITE, m_CREATE_IF_NON_EXISTENT)
			End If
		End If
		Set Init = Me
	End Function

	Private Sub Class_Terminate 'Destructor
    End Sub

	Public Property Get LogFile()
		LogFile = m_strLogFile
	End Property
	
	Public Property Let LogFile(param)
		m_strLogFile = param
	End Property

	Public Property Get Enabled()
		Enabled = m_blnEnabled
	End Property
	
	Public Property Let Enabled(param)
		m_blnEnabled = param
	End Property

	Public Sub LogThis(ByVal strText)
		If (m_blnEnabled) Then
			m_strTextLocal = strText
			Call g_objFunctions.CreatePrintableString(m_strTextLocal)
			m_objLogFile.WriteLine "[" & Now & "]" & "     " & m_strTextLocal
		End If
	End Sub

End Class

Class LoggingAndTracingNoTimestamp
	'
	' Usage:
	'	strTraceFile = "F:\Scripts\CurrentWork\DIGS\DIGS-SQL\VBScriptIncludeCode\junk.txt"
	'	Dim LogAndTrace : Set LogAndTrace = (New LoggingAndTracing)(strTraceFile, True, True)
	'
	'	or
	'
	'	strTraceFile = ""
	'	Dim LogAndTrace : Set LogAndTrace = (New LoggingAndTracing)(strTraceFile, False, True)
	'
	Private m_FOR_WRITE, m_FOR_APPEND, m_CREATE_IF_NON_EXISTENT, m_strLogFile, m_objFSO, m_objLogFile
	Private m_blnEnabled, m_blnCreateNewFile, m_strTextLocal

	Public Default Function Init(LogFile, Enabled, CreateNewFile) 'Constructor
		If (IsObject(g_objFunctions) = False) Then
			WScript.Echo "Object g_objFunctions required for Class LoggingAndTracing.  Abending..."
			WScript.Quit
		End If
		m_FOR_WRITE	= 2
		m_FOR_APPEND = 8
		m_CREATE_IF_NON_EXISTENT = True
		m_strLogFile = LogFile
		m_blnEnabled = Enabled
		m_blnCreateNewFile = CreateNewFile
		If (m_blnEnabled) Then
			Set m_objFSO = CreateObject("Scripting.FileSystemObject")
			If (m_blnCreateNewFile) Then
				If(m_objFSO.FileExists(m_strLogFile)) Then
					m_objFSO.DeleteFile m_strLogFile, True
				End If
				'
				' Open for Write and create if non-existent
				'
				Set m_objLogFile = m_objFSO.OpenTextFile(m_strLogFile, m_FOR_WRITE, m_CREATE_IF_NON_EXISTENT)
			End If
		End If
		Set Init = Me
	End Function

	Private Sub Class_Terminate 'Destructor
    End Sub

	Public Property Get LogFile()
		LogFile = m_strLogFile
	End Property
	
	Public Property Let LogFile(param)
		m_strLogFile = param
	End Property

	Public Property Get Enabled()
		Enabled = m_blnEnabled
	End Property
	
	Public Property Let Enabled(param)
		m_blnEnabled = param
	End Property

	Public Sub LogThis(ByVal strText)
		If (m_blnEnabled) Then
			m_strTextLocal = strText
			Call g_objFunctions.CreatePrintableString(m_strTextLocal)
			m_objLogFile.WriteLine m_strTextLocal
		End If
	End Sub

End Class

Class ClientResolution
	'
	' Requirements:		None
	'
	Private m_adVarChar, m_adStateOpen
	
	Private Sub Class_Initialize() 'Constructor
		If (IsObject(g_objFunctions) = False) Then
			WScript.Echo "Object g_objFunctions required for Class ClientResolution.  Abending..."
			WScript.Quit
		End If
		'
		' ADO Constants
		'
		m_adVarChar = 200
		m_adStateOpen = 1
	End Sub

	Private Sub Class_Terminate 'Destructor
    End Sub
	
	Private Sub LogThis(ByVal strText, ByRef objLogAndTrace)
		Dim strTextLocal
		If (IsObject(objLogAndTrace)) Then
			strTextLocal = strText
			Call g_objFunctions.CreatePrintableString(strTextLocal)
			objLogAndTrace.LogThis(strTextLocal)
		End If
	End Sub

	Public Function GetActiveIPAddress(ByRef objRemoteWMIServer, ByVal intFlag, ByRef strIPAddress, ByRef objLogAndTrace)
	'*****************************************************************************************************************************************
	'*  Purpose:				Gets the IP Address that is currently in use
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				GatherMachineData
	'*  Calls:					ExecWMI, VerifyAndLoad
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim rsGeneric, intRetVal, strTemp, arrTemp
	
		Call LogThis("Processing GetActiveIPAddress started", objLogAndTrace)
		Set rsGeneric = CreateObject("ADODB.Recordset")
		rsGeneric.Fields.Append "SavedData", m_adVarChar, 255
		rsGeneric.Open
		'
		' Get the active interface
		'
		intRetVal = g_objFunctions.ExecCmdGeneric("route print", rsGeneric, "")
		If (intRetVal = 0) Then
			'
			' Got something back
			'
			If (Not rsGeneric.BOF) Then
				rsGeneric.MoveFirst
			End If
			While Not rsGeneric.EOF
				strTemp = rsGeneric("SavedData")
				If (InStr(1, strTemp, "0.0.0.0") > 0) Then
					'
					' IPv4 Route Table
					' ===========================================================================
					' Active Routes:
					' Network Destination        Netmask          Gateway       Interface  Metric
					'           0.0.0.0          0.0.0.0    192.168.182.1   192.168.183.23     30
					'
					strTemp = Trim(strTemp)
'					WScript.Echo "strTemp: " & strTemp
					'0.0.0.0          0.0.0.0    192.168.182.1   192.168.183.23     30
					strTemp = Trim(Split(strTemp, " ", 2)(1))
'					WScript.Echo "strTemp1: " & strTemp
					'0.0.0.0    192.168.182.1   192.168.183.23     30
					strTemp = Trim(Split(strTemp, " ", 2)(1))
'					WScript.Echo "strTemp2: " & strTemp
					'192.168.182.1   192.168.183.23     30
					strTemp = Trim(Split(strTemp, " ", 2)(1))
'					WScript.Echo "strTemp3: " & strTemp
					'192.168.183.23     30
					strTemp = Trim(Split(strTemp, " ", 2)(0))
'					WScript.Echo "strTemp4: " & strTemp
					'192.168.183.23
					strIPAddress = strTemp
					Call LogThis("IP Address: " & strIPAddress, objLogAndTrace)
					Exit Function
				End If
				rsGeneric.MoveNext
			Wend
		End If
		Call LogThis("IP Address not found in route print command", objLogAndTrace)

	End Function

	Private Function CheckClientRPCPING(ByVal strRPCPingFile, ByVal strNameOrIPAddress, ByRef strHostName, ByRef strHostIPAddress, _
											ByRef objLogAndTraceECG)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if a client is alive on network via RPCPING
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 if successful; <> 0 if unsuccessful
	'*  Called by:				IsClientAlive
	'*  Calls:					ExecCmdGeneric, DeleteAllRecordsetRows
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim rsGeneric, strCommand, intRetVal, strRead, intPOS
	
		Set rsGeneric = CreateObject("ADODB.Recordset")
		rsGeneric.Fields.Append "SavedData", m_adVarChar, 255
		rsGeneric.Open

		strCommand = strRPCPingFile &  " -s " & strNameOrIPAddress
		intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, rsGeneric, objLogAndTraceECG)
		If (intRetVal = 0) Then
			'
			' Got something back
			'
			If (Not rsGeneric.BOF) Then
				rsGeneric.MoveFirst
			End If
			Do While Not rsGeneric.EOF
				'
				' Data was returned.  Parse it.
				'
				strRead = rsGeneric("SavedData")
				If (strRead <> "") Then
					intPOS = InStr(1, strRead, "Exception", vbTextCompare)
					If (intPOS > 0) Then
						CheckClientRPCPING = -1
						Call DeleteAllRecordsetRows(rsGeneric)
				 		If (rsGeneric.State = m_adStateOpen) Then
							rsGeneric.Close
						End If
						Set rsGeneric = Nothing
						Exit Function
					End If
					intPOS = InStr(1, strRead, "Completed", vbTextCompare)
					If (intPOS > 0) Then
						CheckClientRPCPING = 0
						Call DeleteAllRecordsetRows(rsGeneric)
				 		If (rsGeneric.State = m_adStateOpen) Then
							rsGeneric.Close
						End If
						Set rsGeneric = Nothing
						Exit Function
					End If
				End If
				rsGeneric.MoveNext
			Loop
		End If
		Call g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
 		If (rsGeneric.State = m_adStateOpen) Then
			rsGeneric.Close
		End If
		Set rsGeneric = Nothing
		CheckClientRPCPING = -1
	
	End Function

	Public Function CheckClientDNS_DNSQuery(ByVal strSiteDCFQDN, ByVal strNameOrIPAddress, ByVal strDomainDNSName, ByRef strHostIPAddress, _
												ByRef objLogAndTraceCR, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if a client is on network via DNS (using NSLookup)
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 if successful, -1 if unsuccessful.
	'*  Called by:				GetClientNetworkInfo
	'*  Calls:					ExecCmdGeneric, DeleteAllRecordsetRows
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim strNameSpaceForConnection, strNameSpaceToValidate, strNameSpace, blnNamespaceInstalled, intRetVal, objRemoteWMIServer, intErrNumber
		Dim strErrDescription, strError, strSQLQuery, intFlag, colWMI, objWMI
		Const wbemFlagForwardOnly = 32
		Const wbemFlagReturnImmediately = 16

		Call LogThis("Executing CheckClientDNS_DNSQuery", objLogAndTraceCR)
		Call LogThis("NameOrIPAddress: " & strNameOrIPAddress, objLogAndTraceCR)
		Call LogThis("DomainDNSName: " & strDomainDNSName, objLogAndTraceCR)

		strHostIPAddress = ""
		strNameSpaceForConnection = "root"
		strNameSpaceToValidate = "MicrosoftDNS"
		strNameSpace = "root\MicrosoftDNS"
		blnNamespaceInstalled = g_objFunctions.IsNamespaceInstalled(strSiteDCFQDN, strNameSpaceForConnection, _
																		strNameSpaceToValidate, "", "", objLogAndTraceErrors)
		If (blnNamespaceInstalled = False) Then
			CheckClientDNS_DNSQuery = -1
			Exit Function
		End If
		intRetVal = g_objFunctions.CreateServerConnection(strSiteDCFQDN, objRemoteWMIServer, intErrNumber, strErrDescription, _
															strError, strNameSpace, "", "", objLogAndTraceErrors)
		If (intRetVal <> 0) Then
			CheckClientDNS_DNSQuery = -1
			Exit Function
		End If
		strSQLQuery = "SELECT * FROM MicrosoftDNS_AType where OwnerName = '" & strNameOrIPAddress & "' and DomainName = '" & _
							strDomainDNSName & "' And ContainerName = '" & strDomainDNSName & "'"
'		intFlag = wbemFlagForwardOnly + wbemFlagReturnImmediately
		intFlag = 0
		Call g_objFunctions.ExecWMI(objRemoteWMIServer, intErrNumber, strErrDescription, colWMI, strSQLQuery, intFlag, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
			For Each objWMI In colWMI
				strHostIPAddress = objWMI.IPAddress
			Next
			CheckClientDNS_DNSQuery = 0
			Exit Function
		End If
		CheckClientDNS_DNSQuery = -1

	End Function

	Public Function CheckClientDNS_NSLookup(ByRef strNameOrIPAddress, ByRef strHostName, ByRef strHostIPAddress, ByRef objLogAndTraceCR, _
												ByRef objLogAndTraceECG)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if a client is on network via DNS (using NSLookup)
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 if successful, -1 if unsuccessful.
	'*  Called by:				GetClientNetworkInfo
	'*  Calls:					ExecCmdGeneric, DeleteAllRecordsetRows
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim rsGeneric, strCommand, intRetVal, strTemp, intPOS, arrTemp, strErrorInfo

		Call LogThis("Executing CheckClientDNS_NSLookup", objLogAndTraceCR)
		Call LogThis("NameOrIPAddress: " & strNameOrIPAddress, objLogAndTraceCR)

		strHostName = ""
		strHostIPAddress = ""

		Set rsGeneric = CreateObject("ADODB.Recordset")
		rsGeneric.Fields.Append "SavedData", m_adVarChar, 255
		rsGeneric.Open

'		strCommand = "nslookup " & strNameOrIPAddress
		strCommand = "nslookup -type=A " & strNameOrIPAddress
		Call LogThis("Calling ExecCmdGeneric from CheckClientDNS_NSLookup", objLogAndTraceCR)
		intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, rsGeneric, objLogAndTraceECG)
		If (intRetVal = 0) Then
			'
			' Got something back
			'
			If (Not rsGeneric.BOF) Then
				rsGeneric.MoveFirst
			End If
			While Not rsGeneric.EOF
				strTemp = rsGeneric("SavedData")
				Call LogThis(strTemp, objLogAndTraceCR)
				If ((InStr(1, strTemp, "***")) Or _
					(InStr(1, strTemp, "dns request timed out", vbTextCompare)) Or _
					(InStr(1, strTemp, "access is denied", vbTextCompare))) Then
					CheckClientDNS_NSLookup = -1
					Call g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
			 		If (rsGeneric.State = m_adStateOpen) Then
						rsGeneric.Close
					End If
					Set rsGeneric = Nothing
					Exit Function
				End If
				If (InStr(1, strTemp, "name:", vbTextCompare)) Then
					arrTemp = Split(strTemp, ":", 2)
					strHostName = Trim(arrTemp(1))
					rsGeneric.MoveNext
					strTemp = rsGeneric("SavedData")
					If (InStr(1, strTemp, "address:", vbTextCompare)) Then
						strHostIPAddress = Trim(Split(strTemp, ":", 2)(1))
						CheckClientDNS_NSLookup = 0
						Call g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
				 		If (rsGeneric.State = m_adStateOpen) Then
							rsGeneric.Close
						End If
						Set rsGeneric = Nothing
						Exit Function
					End If
					'
					' Handle multiple addresses returned by NSLookup
					'
					'	D:\Pub\OSApps_Tools\DataGatheringScript>nslookup fhilapsmsmp01.hill.afmc.ds.af.mil
					'	Server:  fhidc01.hill.afmc.ds.af.mil
					'	Address:  137.241.10.230
					'
					'	Name:    fhilapsmsmp01.hill.afmc.ds.af.mil
					'	Addresses:  137.241.9.122, 137.241.9.123
					'
					If (InStr(1, strTemp, "addresses:", vbTextCompare)) Then
						strHostIPAddress = Trim(Split((Split(strTemp, ":", 2)(1)) , ",", 2)(0))
						CheckClientDNS_NSLookup = 0
						Call g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
				 		If (rsGeneric.State = m_adStateOpen) Then
							rsGeneric.Close
						End If
						Set rsGeneric = Nothing
						Exit Function
					End If
				End If
				rsGeneric.MoveNext
			Wend
		End If
		Call g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
 		If (rsGeneric.State = m_adStateOpen) Then
			rsGeneric.Close
		End If
		Set rsGeneric = Nothing
		CheckClientDNS_NSLookup = -1

	End Function

	Private Function CheckClientPING(ByVal strNameOrIPAddress, ByRef strHostName, ByRef strHostIPAddress, ByRef objLogAndTraceCR, _
										ByRef objLogAndTraceECG)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if a client is alive on network via PING
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 if successful; <> 0 if unsuccessful
	'*  Called by:				GetClientNetworkInfo
	'*  Calls:					ExecCmdGeneric, DeleteAllRecordsetRows
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim rsGeneric, strCommand, intRetVal, strRead, arrRead, strHoldHostName, strHoldHostIPAddress

		Call LogThis("Executing CheckClientPING", objLogAndTraceCR)
		Call LogThis(vbTab & "NameOrIPAddress: " & strNameOrIPAddress, objLogAndTraceCR)

		strHostName = ""
		strHostIPAddress = ""

		Set rsGeneric = CreateObject("ADODB.Recordset")
		rsGeneric.Fields.Append "SavedData", m_adVarChar, 255
		rsGeneric.Open
		'
		' Wait up to 10 seconds for completion of the command
		'
		strCommand = "ping -n 2 -w 10000 -a -4 " & strNameOrIPAddress
		Call LogThis(vbTab & "Calling ExecCmdGeneric from CheckClientPING", objLogAndTraceCR)
		intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, rsGeneric, objLogAndTraceECG)
		If (intRetVal = 0) Then
			'
			' Got something back
			'
			If (Not rsGeneric.BOF) Then
				rsGeneric.MoveFirst
			End If
			While Not rsGeneric.EOF
				strRead = rsGeneric("SavedData")
				Call LogThis(strRead, objLogAndTraceCR)
				If (strRead <> "") Then
					'
					' Process failures first
					'
					If (InStr(1, strRead, "(100% Loss)", vbTextCompare) > 0) Then
						CheckClientPING = -1
						Call g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
				 		If (rsGeneric.State = m_adStateOpen) Then
							rsGeneric.Close
						End If
						Set rsGeneric = Nothing
						Exit Function
					End If
					If (InStr(1, strRead, "ping request could not find host", vbTextCompare) > 0) Then
						CheckClientPING = -1
						Call g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
				 		If (rsGeneric.State = m_adStateOpen) Then
							rsGeneric.Close
						End If
						Set rsGeneric = Nothing
						Exit Function
					End If
					If (InStr(1, strRead, "request timed out", vbTextCompare) > 0) Then
						CheckClientPING = -1
						Call g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
				 		If (rsGeneric.State = m_adStateOpen) Then
							rsGeneric.Close
						End If
						Set rsGeneric = Nothing
						Exit Function
					End If
					'
					' If the Ping is successful (at least partially) then save the
					' host name and IP Address information from the following line
					'
					If (InStr(1, strRead, "Pinging", vbTextCompare) > 0) Then
						If ((InStr(strRead, "[") > 0) And _
							(InStr(strRead, "]") > 0)) Then
							'
							' Resolution with the syntax: "Pinging TLAB-ISO01-N.tlab.centaf.ds.af.mil [153.29.30.90] with 32 bytes of data" occurred
							'
							arrRead = Split(strRead)
							strHoldHostName = Trim(arrRead(1))
							strHoldHostIPAddress = Trim(Replace(Replace(arrRead(2), "[", ""), "]", ""))
							strHostName = strHoldHostName
							strHostIPAddress =  strHoldHostIPAddress
						Else
							'
							' Resolution with the syntax: "Pinging 153.29.22.155 with 32 bytes of data" occurred
							'
							arrRead = Split(strRead)
							strHoldHostName = ""
							strHoldHostIPAddress = Trim(arrRead(1))
						End If
					End If
					If (InStr(1, strRead, "(0% Loss)", vbTextCompare) > 0) Then
						strHostName = strHoldHostName
						strHostIPAddress = strHoldHostIPAddress
						CheckClientPING = 0
						Call g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
				 		If (rsGeneric.State = m_adStateOpen) Then
							rsGeneric.Close
						End If
						Set rsGeneric = Nothing
						Exit Function
					End If
					If (InStr(1, strRead, "(50% Loss)", vbTextCompare) > 0) Then
						strHostName = strHoldHostName
						strHostIPAddress = strHoldHostIPAddress
						CheckClientPING = 0
						Call g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
				 		If (rsGeneric.State = m_adStateOpen) Then
							rsGeneric.Close
						End If
						Set rsGeneric = Nothing
						Exit Function
					End If
				End If
				rsGeneric.MoveNext
			Wend
		End If
		Call g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
 		If (rsGeneric.State = m_adStateOpen) Then
			rsGeneric.Close
		End If
		Set rsGeneric = Nothing
		CheckClientPING = -1

	End Function

	Private Function CheckClientNBTStat(ByVal strNameOrIPAddress, ByRef strHostName, ByRef objLogAndTraceCR, ByRef objLogAndTraceECG)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines the name of a client using NBTStat
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 if successful; <> 0 if unsuccessful
	'*  Called by:				GetClientNetworkInfo
	'*  Calls:					ExecCmdGeneric, DeleteAllRecordsetRow
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim rsGeneric, strCommand, intRetVal, strRead

		strHostName = ""
		'
		' D:\NHA\TLAB\Tools\NA\CheckClientAlive>nbtstat -A 153.29.18.27
		'
		' Team #0 - Adapter Fault Tolerance Mode:
		' Node IpAddress: [153.29.30.90] Scope Id: []
		'
		'           NetBIOS Remote Machine Name Table
		'
		'       Name               Type         Status
		'    ---------------------------------------------
		'    407-EFSS-67763 <00>  UNIQUE      Registered
		'    TLAB-N         <00>  GROUP       Registered
		'    407-EFSS-67763 <20>  UNIQUE      Registered
		'
		'    MAC Address = 00-16-76-A7-92-82
		'
		Call LogThis("Executing CheckClientNBTStat", objLogAndTraceCR)
		Call LogThis(vbTab & "NameOrIPAddress: " & strNameOrIPAddress, objLogAndTraceCR)

		Set rsGeneric = CreateObject("ADODB.Recordset")
		rsGeneric.Fields.Append "SavedData", m_adVarChar, 255
		rsGeneric.Open
		'
		' Wait up to 10 seconds for completion of the command
		'
		strCommand = "nbtstat -A " & strNameOrIPAddress
		Call LogThis(vbTab & "Calling ExecCmdGeneric from CheckClientNBTStat", objLogAndTraceCR)
		intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, rsGeneric, objLogAndTraceECG)
		If (intRetVal = 0) Then
			'
			' Got something back
			'
			If (Not rsGeneric.BOF) Then
				rsGeneric.MoveFirst
			End If
			While Not rsGeneric.EOF
				strRead = rsGeneric("SavedData")
				Call LogThis(strRead, objLogAndTraceCR)
				If (strRead <> "") Then
					If ((InStr(strRead, "<00>", vbTextCompare) > 0) And _
						(InStr(strRead, "Unique", vbTextCompare) > 0)) Then
						strHostName = Split(Trim(strRead), " ")(0)
						CheckClientNBTStat = 0
						Call g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
				 		If (rsGeneric.State = m_adStateOpen) Then
							rsGeneric.Close
						End If
						Set rsGeneric = Nothing
						Exit Function
					End If
				End If
				rsGeneric.MoveNext
			Wend
		End If
		Call g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
 		If (rsGeneric.State = m_adStateOpen) Then
			rsGeneric.Close
		End If
		Set rsGeneric = Nothing
		CheckClientNBTStat = -1

	End Function

	Private Function CheckClientWIN32_PINGSTATUS(ByVal strNameOrIPAddress, ByRef strHostName, ByRef strHostIPAddress, _
													ByRef objLogAndTraceCR, ByRef objLogAndTraceECG)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if a client is alive on network via WIN32_PINGSTATUS
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 if successful; <> 0 if unsuccessful
	'*  Called by:				GetClientNetworkInfo
	'*  Calls:					GetOSVersion, LogThis, ExecWMI
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim objShell, strThisComputer, objWMILocal, strWindowsVersion, strSQLQuery, colPings, intErrNumber, strErrDescription, objPing
		Const wbemFlagReturnWhenComplete = 0
		Const OS_VERSION_XP = "5.1"

		Set objShell = CreateObject("WScript.Shell")
		strThisComputer = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
		Set objWMILocal = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\" & strThisComputer & "\root\cimv2")
		strWindowsVersion = g_objFunctions.GetOSVersion()

		Call LogThis("Executing CheckClientWIN32_PINGSTATUS", objLogAndTraceCR)
		Call LogThis(vbTab & "NameOrIPAddress: " & strNameOrIPAddress, objLogAndTraceCR)
		Call LogThis(vbTab & "Global WindowsVersion: " & strWindowsVersion, objLogAndTraceCR)
		'
		' The Win32_PingStatus is used for pinging, which requires XP or higher. This script can actually be
		' run from a 2000 PC if the ConnectServer string points to an XP box.
		'
		If (strWindowsVersion >= OS_VERSION_XP) Then
			'
			' Processing on a Windows XP or higher machine.  Use the Win32_PingStatus WMI call.
			'
			strSQLQuery = "SELECT StatusCode,PrimaryAddressResolutionStatus,ProtocolAddress,ProtocolAddressResolved " & _
							"FROM Win32_PingStatus WHERE ResolveAddressNames = 'True' AND Address = '" & Replace(strNameOrIPAddress, "'", "''") & "'"
			Call g_objFunctions.ExecWMI(objWMILocal, intErrNumber, strErrDescription, colPings, strSQLQuery, wbemFlagReturnWhenComplete, Null)
			If ((intErrNumber=0) And (UCase(TypeName(colPings))="SWBEMOBJECTSET")) Then
				For Each objPing In colPings
					If ((Not IsNull(objPing.StatusCode)) And (objPing.StatusCode=0) And (objPing.PrimaryAddressResolutionStatus=0)) Then
						strHostIPAddress = objPing.ProtocolAddress
						strHostName = objPing.ProtocolAddressResolved
						'
						' Cleanup
						'
						Set objWMILocal = Nothing
						Set colPings = Nothing
						Set objPing = Nothing
						CheckClientWIN32_PINGSTATUS = 0
						Exit Function
					End If
				Next
			End If
'		Else
'			WScript.Echo objPing.StatusCode
'			Select Case objPing.StatusCode
'				Case 0 CheckClientWIN32_PINGSTATUS = "Success"
'				Case 11001 CheckClientWIN32_PINGSTATUS = "Status code 11001 - Buffer Too Small"
'				Case 11002 CheckClientWIN32_PINGSTATUS = "Status code 11002 - Destination Net Unreachable"
'				Case 11003 CheckClientWIN32_PINGSTATUS = "Status code 11003 - Destination Host Unreachable"
'				Case 11004 CheckClientWIN32_PINGSTATUS = "Status code 11004 - Destination Protocol Unreachable"
'				Case 11005 CheckClientWIN32_PINGSTATUS = "Status code 11005 - Destination Port Unreachable"
'				Case 11006 CheckClientWIN32_PINGSTATUS = "Status code 11006 - No Resources"
'				Case 11007 CheckClientWIN32_PINGSTATUS = "Status code 11007 - Bad Option"
'				Case 11008 CheckClientWIN32_PINGSTATUS = "Status code 11008 - Hardware Error"
'				Case 11009 CheckClientWIN32_PINGSTATUS = "Status code 11009 - Packet Too Big"
'				Case 11010 CheckClientWIN32_PINGSTATUS = "Status code 11010 - Request Timed Out"
'				Case 11011 CheckClientWIN32_PINGSTATUS = "Status code 11011 - Bad Request"
'				Case 11012 CheckClientWIN32_PINGSTATUS = "Status code 11012 - Bad Route"
'				Case 11013 CheckClientWIN32_PINGSTATUS = "Status code 11013 - TimeToLive Expired Transit"
'				Case 11014 CheckClientWIN32_PINGSTATUS = "Status code 11014 - TimeToLive Expired Reassembly"
'				Case 11015 CheckClientWIN32_PINGSTATUS = "Status code 11015 - Parameter Problem"
'				Case 11016 CheckClientWIN32_PINGSTATUS = "Status code 11016 - Source Quench"
'				Case 11017 CheckClientWIN32_PINGSTATUS = "Status code 11017 - Option Too Big"
'				Case 11018 CheckClientWIN32_PINGSTATUS = "Status code 11018 - Bad Destination"
'				Case 11032 CheckClientWIN32_PINGSTATUS = "Status code 11032 - Negotiating IPSEC"
'				Case 11050 CheckClientWIN32_PINGSTATUS = "Status code 11050 - General Failure"
'				Case Else CheckClientWIN32_PINGSTATUS = "Status code " & objPing.StatusCode & " - Unable to determine cause of failure."
'			End Select
		End If
		'
		' Cleanup
		'
		Set objWMILocal = Nothing
		Set colPings = Nothing
		Set objPing = Nothing
		CheckClientWIN32_PINGSTATUS = -1

	End Function

	Public Function ResolveADIZDNSAddress(ByVal strDomainDN, ByVal strDomainDNSName, ByVal strComputer, ByRef strHostIPAddr)
	'*****************************************************************************************************************************************
	'*  Purpose:				Find the IP Address for a client that exists in an Active Directory Integrated DNS Zone
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success, -1 to indicate failure.  IP Address returned in strIPAddr if successful
	'*  Called by:				All
	'*  Calls:					OctetToHexStr, HexToDec
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim objContainer, intErrNumber, strErrDescription, at, strOctet, strDNSData, strTemp, intIndex, strDigit
		' 
		' Don't mess with the next statement - it works (keep as example for correct syntax)
		'
		'''''Set objContainer = GetObject("LDAP://vejxamcw2dc102.amc.ds.af.mil/dc=amc.ds.af.mil,cn=MicrosoftDNS,CN=System,DC=AMC,DC=DS,DC=AF,DC=MIL")
		'
		On Error Resume Next
		Set objContainer = GetObject("LDAP://" & strDomainDNSName & "/dc=" & strDomainDNSName & ",cn=MicrosoftDNS,CN=System," & strDomainDN)
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
		If (intErrNumber  <> 0) Then
			ResolveADIZDNSAddress = -1
			Exit Function
		End If
		On Error Resume Next
		Set at = objContainer.GetObject("dnsNode", "DC=" & strComputer)
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
		If (intErrNumber = 0) Then
			'
			' Leave the following statement in for debugging purposes
			'
			'Call GetADPropertyList(at.ADsPath)
			'
			strOctet = ""
			strDNSData = ""
			strHostIPAddr = ""
			On Error Resume Next
			strDNSData = g_objFunctions.OctetToHexStr(at.dnsRecord)
			On Error GoTo 0
			If (Len(strDNSData) > 0) Then
				strTemp = Right(strDNSData, 8)
				For intIndex = 1 To Len(strTemp) - 1 Step 2
					strDigit = Mid(strTemp, intIndex, 2)
					strOctet = strOctet & g_objFunctions.HexToDec(strDigit) & "."
				Next
				strHostIPAddr = Left(strOctet, Len(strOctet) - 1)
				ResolveADIZDNSAddress = 0
				Exit Function
			End If
		End If
		'
		' Cleanup
		'
		ResolveADIZDNSAddress = -1

	End Function

	Public Function GetLocalClientNetworkInfo(ByRef strResolvedDNSHostName, ByRef strResolvedHostName, ByRef strResolvedIPAddress, _
												ByRef strResolvedNetBIOSName, ByRef strResolutionType, ByRef blnIsAlive, _
												ByRef blnDNSHostNameResolved, ByRef blnHostNameResolved, ByRef blnNetBIOSNameResolved, _
												ByRef blnIPAddressResolved, ByRef blnResolved, ByRef objLogAndTraceCR)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if a client is alive on network and resolves its IP Address and host name.
	'*  Arguments supplied:		NetBIOS name, FQDN, or IP Address
	'*  Return Value:			0 if successful; <> 0 if unsuccessful
	'*  Called by:				Mainline
	'*  Calls:					CheckClientPING, CheckClientDNS
	'*	Requirements:			g_objFunctions
	'*****************************************************************************************************************************************
		Dim objShell, strThisComputer, strThisDNSDomain, strNameSpace, intFlag, intRetVal, objRemoteWMIServer, intErrNumber
		Dim strErrDescription, strError
		Const wbemFlagReturnWhenComplete = 0

		Set objShell = CreateObject("WScript.Shell")
		strThisComputer = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
		strThisDNSDomain = objShell.ExpandEnvironmentStrings("%USERDNSDOMAIN%")
		'
		' ComputerName will ALWAYS be available
		'
		' UserDNSDomain MOST LIKELY WILL NOT be available if the machine is on the domain and we started a remote process.
		'
		strResolvedHostName = strThisComputer
		blnHostNameResolved = True
		strResolvedNetBIOSName = Left(strResolvedHostName, 15)
		blnNetBIOSNameResolved = True
		If (UCase(strThisDNSDomain) = "%USERDNSDOMAIN%") Then
			strResolvedDNSHostName = ""
			blnDNSHostNameResolved = False
		Else
			strResolvedDNSHostName = LCase(strThisComputer & "." & strThisDNSDomain)
			blnDNSHostNameResolved = True
		End If
		strNameSpace = "root\cimv2"
		strResolvedIPAddress = "0.0.0.0"
		intFlag = wbemFlagReturnWhenComplete
		intRetVal = g_objFunctions.CreateServerConnection(strThisComputer, objRemoteWMIServer, intErrNumber, strErrDescription, _
															strError, strNameSpace, "", "", "")
		If (intRetVal = 0) Then
			Call GetActiveIPAddress(objRemoteWMIServer, intFlag, strResolvedIPAddress, objLogAndTraceCR)
		End If
		If (strResolvedIPAddress = "0.0.0.0") Then
			blnIPAddressResolved = False
		Else
			blnIPAddressResolved = True
		End If
		strResolutionType = "PING"
		blnIsAlive = True
		blnResolved = True
		Set objRemoteWMIServer = Nothing
	
	End Function

	Public Function GetClientNetworkInfo(ByVal strComputerOrIPAddress, ByVal blnPassedDNSHostName, ByVal blnPassedHostName, _
											ByVal blnPassedIPAddress, ByRef strResolvedDNSHostName, ByRef strResolvedHostName, _
											ByRef strResolvedIPAddress, ByRef strResolvedNetBIOSName, ByRef strResolutionType, _
											ByRef blnIsAlive, ByRef blnDNSHostNameResolved, ByRef blnHostNameResolved, _
											ByRef blnNetBIOSNameResolved, ByRef blnIPAddressResolved, ByRef blnResolved, _
											ByRef objLogAndTraceCR, ByRef objLogAndTraceECG)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if a client is alive on network and resolves its IP Address and host name.
	'*  Arguments supplied:		NetBIOS name, FQDN, or IP Address
	'*  Return Value:			0 if successful; <> 0 if unsuccessful
	'*  Called by:				Mainline
	'*  Calls:					CheckClientPING, CheckClientDNS
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim intRetValDNSHostName, strSaveComputerName, strSaveIPAddress, strHostName, intRetValHostName, strNetBIOSName
		Dim intRetValNetBIOSName, intRetValClientDNS, intRetValClientPing

		strResolvedDNSHostName = Empty
		strResolvedHostName = Empty
		strResolvedNetBIOSName = Empty
		strResolvedIPAddress = "0.0.0.0"
		strResolutionType = "Unavailable"
		blnIsAlive = False
		blnDNSHostNameResolved = False
		blnHostNameResolved = False
		blnNetBIOSNameResolved = False
		blnIPAddressResolved = False
		blnResolved = False

		If (blnPassedDNSHostName) Then
			Call LogThis("Passed DNSHostName", objLogAndTraceCR)
			Call LogThis("Attempting resolution using: " & strComputerOrIPAddress, objLogAndTraceCR)
			intRetValDNSHostName = CheckClientPING(strComputerOrIPAddress, strSaveComputerName, strSaveIPAddress, _
														objLogAndTraceCR, objLogAndTraceECG)
			If (intRetValDNSHostName = 0) Then
				Call LogThis(vbTab & "intRetValDNSHostName = 0", objLogAndTraceCR)
				Call LogThis(vbTab & "SaveComputerName: " & strSaveComputerName, objLogAndTraceCR)
				Call LogThis(vbTab & "SaveIPAddress: " & strSaveIPAddress, objLogAndTraceCR)
				strResolvedDNSHostName = strComputerOrIPAddress
				blnDNSHostNameResolved = True
				strResolvedIPAddress = strSaveIPAddress
				blnIPAddressResolved = True
				'
				' We know this is an FQDN - split it to get the host name
				'
				strResolvedHostName = Trim(Split(strSaveComputerName, ".", 2)(0))
				blnHostNameResolved = True
				strResolvedNetBIOSName = Left(strResolvedHostName, 15)
				blnNetBIOSNameResolved = True
				strResolutionType = "PING"
				blnIsAlive = True
				blnResolved = True
			 	GetClientNetworkInfo = 0
				Exit Function
			End If
			'
			' Attempt to PING just the HostName
			'
			strHostName = Split(strComputerOrIPAddress, ".", 2)(0)
			Call LogThis("Attempting resolution using: " & strHostName, objLogAndTraceCR)
			intRetValHostName = CheckClientPING(strHostName, strSaveComputerName, strSaveIPAddress, _
													objLogAndTraceCR, objLogAndTraceECG)
			If (intRetValHostName = 0) Then
				If (InStr(strSaveComputerName, ".") > 0) Then
					strResolvedDNSHostName = strSaveComputerName
					blnDNSHostNameResolved = True
				End If
				strResolvedHostName = strHostName
				blnHostNameResolved = True
				strResolvedNetBIOSName = Left(strResolvedHostName, 15)
				blnNetBIOSNameResolved = True
				strResolvedIPAddress = strSaveIPAddress
				blnIPAddressResolved = True
				If (strResolutionType = "Unavailable") Then
					strResolutionType = "PING"
				End If
				blnIsAlive = True
				blnResolved = True
			 	GetClientNetworkInfo = 0
				Exit Function				
			End If
			'
			' Attempt to PING just the NetBIOS Name (if it is different from the HostName)
			'
			If (Len(strHostName) > 15) Then
				strNetBIOSName = Left(strHostName, 15)
				Call LogThis("Attempting resolution using: " & strNetBIOSName, objLogAndTraceCR)
				intRetValNetBIOSName = CheckClientPING(strNetBIOSName, strSaveComputerName, strSaveIPAddress, _
															objLogAndTraceCR, objLogAndTraceECG)
				If (intRetValNetBIOSName = 0) Then
					Call LogThis(vbTab & "intRetValNetBIOSName = 0", objLogAndTraceCR)
					Call LogThis(vbTab & "SaveComputerName: " & strSaveComputerName, objLogAndTraceCR)
					Call LogThis(vbTab & "SaveIPAddress: " & strSaveIPAddress, objLogAndTraceCR)
					If (InStr(strSaveComputerName, ".") > 0) Then
						strResolvedDNSHostName = strSaveComputerName
						blnDNSHostNameResolved = True
						'
						' We know this is an FQDN - split it to get the host name
						'
						strResolvedHostName = Trim(Split(strSaveComputerName, ".", 2)(0))
						blnHostNameResolved = True
					End If
					strResolvedNetBIOSName = strNetBIOSName
					blnNetBIOSNameResolved = True
					strResolvedIPAddress = strSaveIPAddress
					blnIPAddressResolved = True
					If (strResolutionType = "Unavailable") Then
						strResolutionType = "PING"
					End If
					blnIsAlive = True
					blnResolved = True
				 	GetClientNetworkInfo = 0
					Exit Function				
				End If
			End If
			'
			' Ping didn't do it - see if we can at least get the address from DNS
			'
			Call LogThis("Attempting DNS NSLookup resolution using: " & strComputerOrIPAddress, objLogAndTraceCR)
			intRetValClientDNS = CheckClientDNS_NSLookup(strComputerOrIPAddress, strSaveComputerName, strSaveIPAddress, _
															objLogAndTraceCR, objLogAndTraceECG)
			If (intRetValClientDNS = 0) Then
				Call LogThis(vbTab & "intRetValClientDNS = 0", objLogAndTraceCR)
				Call LogThis(vbTab & "SaveComputerName: " & strSaveComputerName, objLogAndTraceCR)
				Call LogThis(vbTab & "SaveIPAddress: " & strSaveIPAddress, objLogAndTraceCR)
				If (InStr(strSaveComputerName, ".") > 0) Then
					strResolvedDNSHostName = strSaveComputerName
					blnDNSHostNameResolved = True
					'
					' We know this is an FQDN - split it to get the host name
					'
					strResolvedHostName = Trim(Split(strSaveComputerName, ".", 2)(0))
					blnHostNameResolved = True
					strResolvedNetBIOSName = Left(strResolvedHostName, 15)
					blnNetBIOSNameResolved = True
				ElseIf (strSaveComputerName <> "") Then
					strResolvedHostName = strSaveComputerName
					blnHostNameResolved = True
					strResolvedNetBIOSName = Left(strResolvedHostName, 15)
					blnNetBIOSNameResolved = True
				End If
				strResolvedIPAddress = strSaveIPAddress
				blnIPAddressResolved = True
				strResolutionType = "DNS (via NSLOOKUP)"
				blnResolved = True
			 	GetClientNetworkInfo = 0
			 	Exit Function
			End If
			GetClientNetworkInfo = -1
			Exit Function
		End If

		If (blnPassedHostName) Then
			Call LogThis("Passed HostName", objLogAndTraceCR)
			Call LogThis("Attempting resolution using: " & strComputerOrIPAddress, objLogAndTraceCR)
			intRetValHostName = CheckClientPING(strComputerOrIPAddress, strSaveComputerName, strSaveIPAddress, _
													objLogAndTraceCR, objLogAndTraceECG)
			If (intRetValHostName = 0) Then
				Call LogThis(vbTab & "intRetValHostName = 0", objLogAndTraceCR)
				Call LogThis(vbTab & "SaveComputerName: " & strSaveComputerName, objLogAndTraceCR)
				Call LogThis(vbTab & "SaveIPAddress: " & strSaveIPAddress, objLogAndTraceCR)
				If (InStr(strSaveComputerName, ".") > 0) Then
					strResolvedDNSHostName = strSaveComputerName
					blnDNSHostNameResolved = True
				End If
				strResolvedHostName = strComputerOrIPAddress
				blnHostNameResolved = True
				strResolvedNetBIOSName = Left(strResolvedHostName, 15)
				blnNetBIOSNameResolved = True
				strResolvedIPAddress = strSaveIPAddress
				blnIPAddressResolved = True
				strResolutionType = "PING"
				blnIsAlive = True
				blnResolved = True
			 	GetClientNetworkInfo = 0
			 	Exit Function
			End If
			'
			' Attempt to PING just the NetBIOS Name (if it is different from the HostName)
			'
			If (Len(strComputerOrIPAddress) > 15) Then
				strNetBIOSName = Left(strComputerOrIPAddress, 15)
				Call LogThis("Attempting resolution using: " & strNetBIOSName, objLogAndTraceCR)
				intRetValNetBIOSName = CheckClientPING(strNetBIOSName, strSaveComputerName, strSaveIPAddress, _
															objLogAndTraceCR, objLogAndTraceECG)
				If (intRetValNetBIOSName = 0) Then
					Call LogThis(vbTab & "intRetValNetBIOSName = 0", objLogAndTraceCR)
					Call LogThis(vbTab & "SaveComputerName: " & strSaveComputerName, objLogAndTraceCR)
					Call LogThis(vbTab & "SaveIPAddress: " & strSaveIPAddress, objLogAndTraceCR)
					If (InStr(strSaveComputerName, ".") > 0) Then
						strResolvedDNSHostName = strSaveComputerName
						blnDNSHostNameResolved = True
						'
						' We know this is an FQDN - split it to get the host name
						'
						strResolvedHostName = Trim(Split(strSaveComputerName, ".", 2)(0))
						blnHostNameResolved = True
					End If
					strResolvedNetBIOSName = strNetBIOSName
					blnNetBIOSNameResolved = True
					strResolvedIPAddress = strSaveIPAddress
					blnIPAddressResolved = True
					If (strResolutionType = "Unavailable") Then
						strResolutionType = "PING"
					End If
					blnIsAlive = True
					blnResolved = True
				 	GetClientNetworkInfo = 0
				 	Exit Function
				End If
			End If
			'
			' Ping didn't do it - see if we can at least get the address from DNS
			'
			Call LogThis("Attempting DNS NSLookup resolution using: " & strComputerOrIPAddress, objLogAndTraceCR)
			intRetValClientDNS = CheckClientDNS_NSLookup(strComputerOrIPAddress, strSaveComputerName, strSaveIPAddress, _
															objLogAndTraceCR, objLogAndTraceECG)
			If (intRetValClientDNS = 0) Then
				Call LogThis(vbTab & "intRetValClientDNS = 0", objLogAndTraceCR)
				Call LogThis(vbTab & "SaveComputerName: " & strSaveComputerName, objLogAndTraceCR)
				Call LogThis(vbTab & "SaveIPAddress: " & strSaveIPAddress, objLogAndTraceCR)
				If (InStr(strSaveComputerName, ".") > 0) Then
					strResolvedDNSHostName = strSaveComputerName
					blnDNSHostNameResolved = True
					'
					' We know this is an FQDN - split it to get the host name
					'
					strResolvedHostName = Trim(Split(strSaveComputerName, ".", 2)(0))
					blnHostNameResolved = True
					strResolvedNetBIOSName = Left(strResolvedHostName, 15)
					blnNetBIOSNameResolved = True
				ElseIf (strSaveComputerName <> "") Then
					strResolvedHostName = strSaveComputerName
					blnHostNameResolved = True
					strResolvedNetBIOSName = Left(strResolvedHostName, 15)
					blnNetBIOSNameResolved = True
				End If
				strResolvedIPAddress = strSaveIPAddress
				blnIPAddressResolved = True
				strResolutionType = "DNS (via NSLOOKUP)"
				blnResolved = True
			 	GetClientNetworkInfo = 0
			 	Exit Function
			End If
			GetClientNetworkInfo = -1
			Exit Function
		End If

		If (blnPassedIPAddress) Then
			Call LogThis("Passed IPAddress", objLogAndTraceCR)
			Call LogThis("Attempting resolution using: " & strComputerOrIPAddress, objLogAndTraceCR)
			intRetValClientPing = CheckClientPING(strComputerOrIPAddress, strSaveComputerName, strSaveIPAddress, _
														objLogAndTraceCR, objLogAndTraceECG)
			If (intRetValClientPing = 0) Then
				Call LogThis(vbTab & "intRetValClientPing = 0", objLogAndTraceCR)
				Call LogThis(vbTab & "SaveComputerName: " & strSaveComputerName, objLogAndTraceCR)
				Call LogThis(vbTab & "SaveIPAddress: " & strSaveIPAddress, objLogAndTraceCR)
				If (InStr(strSaveComputerName, ".") > 0) Then
					strResolvedDNSHostName = strSaveComputerName
					blnDNSHostNameResolved = True
					'
					' We know this is an FQDN - split it to get the host name
					'
					strResolvedHostName = Trim(Split(strSaveComputerName, ".", 2)(0))
					blnHostNameResolved = True
					strResolvedNetBIOSName = Left(strResolvedHostName, 15)
					blnNetBIOSNameResolved = True
				ElseIf (strSaveComputerName <> "") Then
					strResolvedHostName = strSaveComputerName
					blnHostNameResolved = True
					strResolvedNetBIOSName = Left(strResolvedHostName, 15)
					blnNetBIOSNameResolved = True
				End If
				strResolvedIPAddress = strSaveIPAddress
				blnIPAddressResolved = True
				strResolutionType = "PING"
				blnIsAlive = True
				blnResolved = True
			 	GetClientNetworkInfo = 0
			 	Exit Function
			End If
			'
			' Ping didn't do it - see if we can at least get the address from DNS
			'
			Call LogThis("Attempting DNS NSLookup resolution using: " & strComputerOrIPAddress, objLogAndTraceCR)
			intRetValClientDNS = CheckClientDNS_NSLookup(strComputerOrIPAddress, strSaveComputerName, strSaveIPAddress, _
															objLogAndTraceCR, objLogAndTraceECG)
			If (intRetValClientDNS = 0) Then
				Call LogThis(vbTab & "intRetValClientDNS = 0", objLogAndTraceCR)
				Call LogThis(vbTab & "SaveComputerName: " & strSaveComputerName, objLogAndTraceCR)
				Call LogThis(vbTab & "SaveIPAddress: " & strSaveIPAddress, objLogAndTraceCR)
				If (InStr(strSaveComputerName, ".") > 0) Then
					strResolvedDNSHostName = strSaveComputerName
					blnDNSHostNameResolved = True
					'
					' We know this is an FQDN - split it to get the host name
					'
					strResolvedHostName = Trim(Split(strSaveComputerName, ".", 2)(0))
					blnHostNameResolved = True
					strResolvedNetBIOSName = Left(strResolvedHostName, 15)
					blnNetBIOSNameResolved = True
				ElseIf (strSaveComputerName <> "") Then
					strResolvedHostName = strSaveComputerName
					blnHostNameResolved = True
					strResolvedNetBIOSName = Left(strResolvedHostName, 15)
					blnNetBIOSNameResolved = True
				End If
				strResolvedIPAddress = strSaveIPAddress
				blnIPAddressResolved = True
				strResolutionType = "DNS (via NSLOOKUP)"
				blnResolved = True
			 	GetClientNetworkInfo = 0
			 	Exit Function
			End If
			GetClientNetworkInfo = -1
			Exit Function
		End If
		GetClientNetworkInfo = -1

	End Function

	Public Function GetIPAddressFromHostName(ByVal strSiteDCFQDN, ByVal strDomainDNSName, ByVal strHostName, ByRef strHostIPAddress, _
												ByRef strResolutionType, ByRef objLogAndTraceCR, ByRef objLogAndTraceECG, _
												ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if a client is alive on network and resolves its IP Address.
	'*  Arguments supplied:		NetBIOS name or FQDN
	'*  Return Value:			0 if successful; -1 if unsuccessful
	'*  Called by:				Main()
	'*  Calls:					ResolveADIZDNSAddress, CheckClientWIN32_PINGSTATUS, CheckClientDNS_NSLookup, and CheckClientPING 
	'*****************************************************************************************************************************************
		Dim intRetVal, strRetHostName, blnPartialResolution

' 		intRetVal = ResolveADIZDNSAddress(strDomainDN, strDomainDNSName, strADIZName, strHostIPAddress)
' 		If (intRetVal = 0) Then
' 			strResolutionType = "ADIZ-DNS"
' 			GetIPAddressFromHostName = intRetVal
' 			Exit Function
' 		End If

		intRetVal = CheckClientDNS_NSLookup(strHostName, strRetHostName, strHostIPAddress, objLogAndTraceCR, objLogAndTraceECG)
		If (intRetVal = 0) Then
			strResolutionType = "DNS (via NSLOOKUP)"
			GetIPAddressFromHostName = intRetVal
			Exit Function
		End If

		intRetVal = CheckClientDNS_DNSQuery(strSiteDCFQDN, strHostName, strDomainDNSName, strHostIPAddress, objLogAndTraceCR, _
												objLogAndTraceErrors)
		If (intRetVal = 0) Then
			strResolutionType = "DNS (via MicrosoftDNS_AType)"
			GetIPAddressFromHostName = intRetVal
			Exit Function
		End If
		strHostIPAddress = "0.0.0.0"
		strResolutionType = "Unavailable"
		GetIPAddressFromHostName = -1

	End Function

	Public Function IsClientAlive(ByVal strMachineNameOrIPAddress, ByRef objLogAndTraceCR, ByRef objLogAndTraceECG)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determines if a client is alive on network
	'*  Arguments supplied:		NetBIOS name, FQDN, or IP Address
	'*  Return Value:			0 if successful; <> 0 if unsuccessful
	'*  Called by:				External Programs
	'*  Calls:					DetermineNameOrIPAddress, CheckClientWIN32_PINGSTATUS, and CheckClientPING
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim objShell, rsToProcess, strToUse, blnFQDN, blnMachineName, blnIPAddress, strMachineName, strIPAddress, intRetVal
		Dim strHostName, strHostIPAddress, strToProcess, blnClientAlive

		Set objShell = CreateObject("WScript.Shell")
		Set rsToProcess = CreateObject("ADODB.Recordset")
		rsToProcess.Fields.Append "ToUse", m_adVarChar, 80
		rsToProcess.Open

		If (strMachineNameOrIPAddress = ".") Then
			strToUse = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
		Else
			strToUse = strMachineNameOrIPAddress
		End If

		Call g_objFunctions.DetermineNameOrIPAddress(strToUse, blnFQDN, blnMachineName, blnIPAddress, strMachineName, strIPAddress)
		If (blnFQDN) Then
			rsToProcess.AddNew
			rsToProcess("ToUse") = strToUse
			rsToProcess.Update
		End If
		If (strMachineName <> "") Then
			'
			' strMachineName holds the actual ComputerName extracted from the FQDN
			'
			If (Len(strMachineName) > 15) Then
				rsToProcess.AddNew
				rsToProcess("ToUse") = Left(strMachineName, 15)
				rsToProcess.Update
			Else
				rsToProcess.AddNew
				rsToProcess("ToUse") = strMachineName
				rsToProcess.Update
			End If
		End If
		If (blnIPAddress) Then
			rsToProcess.AddNew
			rsToProcess("ToUse") = strToUse
			rsToProcess.Update
		End If
		'
		' Make sure there is something to process
		'
		If (rsToProcess.RecordCount = 0) Then
			IsClientAlive = False
			Exit Function
		End If
		'
		' Try all combinations that we have to see which one works
		'
		If (Not rsToProcess.BOF) Then
			rsToProcess.MoveFirst
		End If
		While Not rsToProcess.EOF
			strToProcess = rsToProcess("ToUse")
			blnClientAlive = False
			intRetVal = CheckClientWIN32_PINGSTATUS(strToProcess, strHostName, strHostIPAddress, objLogAndTraceCR, objLogAndTraceECG)
			If (intRetVal = 0) Then
				blnClientAlive = True
				If (InStr(strHostIPAddress, ":") = 0) Then
					IsClientAlive = True
					Exit Function
				End If
			End If
			intRetVal = CheckClientPING(strToProcess, strHostName, strHostIPAddress, objLogAndTraceCR, objLogAndTraceECG)
			If (intRetVal = 0) Then
				blnClientAlive = True
				If (strHostName <> "") Then
					IsClientAlive = True
					Exit Function
				End If
			End If
			rsToProcess.MoveNext
		Wend
		IsClientAlive = False
		'
		' Cleanup
		'
		Set objShell = Nothing
		Set rsToProcess = Nothing

	End Function

	Public Function GetMyIPAddress(ByVal strToProcess, ByRef strIPAddress, ByRef objLogAndTraceCR, ByRef objLogAndTraceECG)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get the IP Address of the host computer
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 if successful; <> 0 if unsuccessful
	'*  Called by:				External Programs
	'*  Calls:					DetermineNameOrIPAddress, CheckClientWIN32_PINGSTATUS, and CheckClientPING
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim objShell, strToUse, intRetVal, strHostName

		Set objShell = CreateObject("WScript.Shell")
		If (strToProcess = ".") Then
			strToUse = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
		Else
			strToUse = strToProcess
		End If
		Set objShell = Nothing

		intRetVal = CheckClientWIN32_PINGSTATUS(strToUse, strHostName, strIPAddress, objLogAndTraceCR, objLogAndTraceECG)
		If (intRetVal = 0) Then
			Exit Function
		End If

		intRetVal = CheckClientPING(strToUse, strHostName, strIPAddress, objLogAndTraceCR, objLogAndTraceECG)
		If (intRetVal = 0) Then
			Exit Function
		End If

	End Function

End Class

Class RegistryProcessing
	'
	' Requirements:		None
	'
	'
	' WMI processing functions:
	'
	'	EnumWMIRegistryProcessing
	'		Enumerates all registry keys/values (via WMI) and loads the results in the m_rsRegistryEnum recordset
	'	ExactEntryRegistryProcessing
	'		Searches registry (via WMI) for a specific registry entry
	'	EnumWMIRegistryEntries
	'		Calls EnumWMIRegistryProcessing and ExactEntryRegistryProcessing to get all Entries for the Key passed
	'	GetWMIRegistryDefaultValue
	'		Looks for the (Default) entry.
	'	SetWMIRegistryEntry
	'		Sets the registry entry to the specified value (if the entry doesn't exist it will be created)
	'	DeleteWMIRegistryEntry
	'		Deletes the specified registry entry
	'	DeleteWMIRegistryKey
	'		Deletes the specified registry key
	'
	' Remote Registry processing functions:
	'
	'	EnumRemoteRegistryKeys
	'		Enumerates all registry keys (via RemoteRegistry) using "REG QUERY" and ParseRemoteRegistryKeys and loads the results 
	'		in the m_rsKeys recordset
	'	EnumRemoteRegistryEntries
	'		Enumerates all registry keys (via RemoteRegistry) using "REG QUERY" and ParseRemoteRegistryEntries and loads the results 
	'		in the m_rsEntries recordset
	'	SetRemoteRegistryEntry
	'		Sets the registry entry to the specified value (if the entry doesn't exist it will be created)
	'	DeleteRemoteRegistryEntry
	'		Deletes the specified registry entry
	'	DeleteRemoteRegistryKey
	'		Deletes the specified registry key
	'
	'
	' Global (Public) functions:
	'
	'	RegKeyEntryExists
	'		Checks to see if any registry Entries exists
	'	RegKeySubkeyExists
	'		Checks to see if any registry Subkeys exists
	'	GetHiveInformation
	'		Uses the Registry Hive information (expanded) to return the Registry Hive (Hex) and Registry Hive (short)
	'	GetExpandedHiveInformation
	'		Uses the Registry Hive information (short) to return the Registry Hive (Hex) and Registry Hive (expanded)
	'	EnumRegistryKeys
	'		Calls EnumWMIRegistryProcessing or EnumRemoteRegistryKeys based on access to WMI Registry and/or RemoteRegistry.
	'			A second call may be necessary based on RequireTwoRegistryReads return value.
	'	GetRegistryKey
	'		Calls EnumWMIRegistryProcessing or EnumRemoteRegistryKeys based on access to WMI Registry and/or RemoteRegistry.
	'			A second call may be necessary based on RequireTwoRegistryReads return value.
	'	EnumRegistryEntries
	'		Calls EnumWMIRegistryEntries or EnumRemoteRegistryEntries based on access to WMI Registry and/or RemoteRegistry.
	'			A second call may be necessary based on RequireTwoRegistryReads return value.
	'	GetRegistryEntry
	'		Calls EnumWMIRegistryEntries or EnumRemoteRegistryEntries based on access to WMI Registry and/or RemoteRegistry.
	'			A second call may be necessary based on RequireTwoRegistryReads return value.
	'	GetSpecificRegistryEntry
	'		Calls EnumWMIRegistryEntries or EnumRemoteRegistryEntries based on access to WMI Registry and/or RemoteRegistry.
	'	EnumRegistryEntriesByKey
	'		Calls EnumWMIRegistryEntries or EnumRemoteRegistryEntries based on access to WMI Registry and/or RemoteRegistry.
	'	EnumRegistrySubkeysByKey
	'		Calls EnumWMIRegistryProcessing or EnumRemoteRegistryKeys based on access to WMI Registry and/or RemoteRegistry.
	'	SetRegistryEntry
	'		Calls SetWMIRegistryEntry or SetRemoteRegistryEntry based on access to WMI Registry and/or RemoteRegistry.
	'	DeleteRegistryEntry
	'		Calls DeleteWMIRegistryEntry or DeleteRemoteRegistryEntry based on access to WMI Registry and/or RemoteRegistry.
	'	DeleteRegistryKey
	'		Calls DeleteWMIRegistryKey or DeleteRemoteRegistryKey based on access to WMI Registry and/or RemoteRegistry.
	'
 	Private m_adVarChar, m_adLongVarWChar, m_adInteger, m_adBoolean, m_adStateOpen, m_HKEY_CLASSES_ROOT, m_HKEY_CURRENT_USER
 	Private m_HKEY_LOCAL_MACHINE, m_HKEY_USERS, m_HKEY_CURRENT_CONFIG, m_REG_KEY_QUERY, m_REG_SZ, m_REG_EXPAND_SZ, m_REG_BINARY
 	Private m_REG_DWORD, m_REG_MULTI_SZ, m_rsEntries, m_rsKeys, m_rsRegistryEnum, m_rsRegistryEntries, m_objLogAndTrace
 	Private m_objLogAndTraceErrors, m_objLogAndTraceLoadRS

	Private Sub Class_Initialize() 'Constructor
		If (IsObject(g_objFunctions) = False) Then
			WScript.Echo "Object g_objFunctions required for Class RegistryProcessing.  Abending..."
			WScript.Quit
		End If
		'
		' ADO Constants
		'
		m_adVarChar = 200
		m_adLongVarWChar = 203
		m_adInteger = 3
		m_adBoolean = 11
		m_adStateOpen = 1
		'
		' Registry Constants
		'
		m_HKEY_CLASSES_ROOT = &H80000000
		m_HKEY_CURRENT_USER = &H80000001
		m_HKEY_LOCAL_MACHINE = &H80000002
		m_HKEY_USERS = &H80000003
		m_HKEY_CURRENT_CONFIG = &H80000005
		'
		' CheckAccess Registry constants
		'
		m_REG_KEY_QUERY = &H0001
		'
		' Registry Data Type constants
		'
		m_REG_SZ = 1
		m_REG_EXPAND_SZ = 2
		m_REG_BINARY = 3
		m_REG_DWORD = 4
		m_REG_MULTI_SZ = 7
		'
		' Create m_rsRemoteRegistryEntries recordset
		'
		Set m_rsEntries = CreateObject("ADODB.Recordset")
		m_rsEntries.Fields.Append "Hive", m_adVarChar, 10
		m_rsEntries.Fields.Append "Key", m_adVarChar, 255
		m_rsEntries.Fields.Append "Entry", m_adVarChar, 255
		m_rsEntries.Fields.Append "Type", m_adInteger
		m_rsEntries.Fields.Append "TypeText", m_adVarChar, 20
		m_rsEntries.Fields.Append "Wow6432Node", m_adBoolean
		m_rsEntries.Fields.Append "Data", m_adLongVarWChar, 512
		m_rsEntries.Open
		'
		' Create m_rsRemoteRegistryKeys recordset
		'
		Set m_rsKeys = CreateObject("ADODB.Recordset")
		m_rsKeys.Fields.Append "Hive", m_adVarChar, 10
		m_rsKeys.Fields.Append "Key", m_adVarChar, 255
		m_rsKeys.Fields.Append "Subkey", m_adVarChar, 255
		m_rsKeys.Fields.Append "Wow6432Node", m_adBoolean
		m_rsKeys.Open
		'
		' Create m_rsRegistryEnum recordset
		'
		Set m_rsRegistryEnum = CreateObject("ADODB.Recordset")
		m_rsRegistryEnum.Fields.Append "Hive", m_adVarChar, 10
		m_rsRegistryEnum.Fields.Append "Key", m_adVarChar, 255
		m_rsRegistryEnum.Fields.Append "Subkey", m_adVarChar, 255
		m_rsRegistryEnum.Fields.Append "Type", m_adInteger
		m_rsRegistryEnum.Open
		'
		' Create m_rsRegistryEntries recordset
		'
		Set m_rsRegistryEntries = CreateObject("ADODB.Recordset")
		m_rsRegistryEntries.Fields.Append "Hive", m_adVarChar, 10
		m_rsRegistryEntries.Fields.Append "Key", m_adVarChar, 255
		m_rsRegistryEntries.Fields.Append "Entry", m_adVarChar, 255
		m_rsRegistryEntries.Fields.Append "Type", m_adInteger
		m_rsRegistryEntries.Fields.Append "TypeText", m_adVarChar, 20
		m_rsRegistryEntries.Fields.Append "Data", m_adLongVarWChar, 512
		m_rsRegistryEntries.Open

	End Sub

	Private Sub Class_Terminate 'Destructor
		Set m_rsEntries = Nothing
		Set m_rsKeys = Nothing
		Set m_rsRegistryEnum = Nothing
 		Set m_rsRegistryEntries = Nothing
   End Sub
	
	Private Sub LogThis(ByVal strText, ByRef objLogAndTrace)
		Dim strTextLocal
		If (IsObject(objLogAndTrace)) Then
			strTextLocal = strText
			Call g_objFunctions.CreatePrintableString(strTextLocal)
			objLogAndTrace.LogThis(strTextLocal)
		End If
	End Sub

	Private Function RequireTwoRegistryReads(ByVal strRegPath, ByVal strOSVersion)
	'*****************************************************************************************************************************************
	'*  Purpose:				Checks to see if a 64-bit machine requires the 64-bit and 32-bit hives to be read.
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate single read required
	'*  Called by:				EnumRegistryKeys, EnumRegistryEntries, GetRegistryKey, GetRegistryEntry
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		'
		' http://msdn.microsoft.com/en-us/library/windows/desktop/aa384253(v=vs.85).aspx
		'
		' If the registry path is in the list below then the information in the 64-bit and 32-bit registry locations
		' is shared (contain the same information).  Only read the registry once using standard regisry provider if 
		' there is a match with the search string.
		'
		' Assume that we need to read both
		'
		RequireTwoRegistryReads = True
		'
		' The following are the same regardless of OS Version
		'
		If ((InStr(1, strRegPath, "HKLM\Software\Clients", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Cryptography\Calais\Current", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Cryptography\Calais\Readers", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Cryptography\Services", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\CTF\SystemShared", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\CTF\TIP", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\DFS", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Driver Signing", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\EnterpriseCertificates", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\MSMQ", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Non-Driver Signing", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\RAS", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\SOFTWARE\Microsoft\Shared Tools\MSInfo", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\SystemCertificates", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\TermServLicensing", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\TransactionServer", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows\CurrentVersion\Control Panel\Cursors\Schemes", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows\CurrentVersion\Group Policy", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "Software\Microsoft\Windows\CurrentVersion\Policies", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows\CurrentVersion\Setup", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows\CurrentVersion\Telephony\Locations", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\FontDpi", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\FontMapper", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Fonts", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\FontSubstitutes", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\NetworkCards", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Perflib", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Ports", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Print", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\ProfileList", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Time Zones", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\Policies", vbTextCompare) > 0) Or _
			(InStr(1, strRegPath, "HKLM\Software\RegisteredApplications", vbTextCompare) > 0)) Then 
'			(InStr(1, strRegPath, "HKLM\Software\Network Associates\ePolicy Orchestrator\Application Plugins", vbTextCompare) > 0)) Then
			RequireTwoRegistryReads = False
		End If
		'
		' Check the following for Windows 7 and Windows Server 2008 R2 only
		'
		If (strOSVersion = "6.1") Then
			If ((InStr(1, strRegPath, "HKLM\Software\Clients", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\COM3", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\EventSystem", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\Notepad\DefaultFonts", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\OLE", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\RPC", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows\CurrentVersion\Explorer\KindMap", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows\CurrentVersion\PreviewHandlers", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Console", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\FontLink", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Gre_Initialize", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options", vbTextCompare) > 0) Or _
				(InStr(1, strRegPath, "HKLM\Software\Microsoft\Windows NT\CurrentVersion\Language Pack", vbTextCompare) > 0)) Then
				RequireTwoRegistryReads = False
			End If
		End If

	End Function

'#region <WMI Registry Processing Functions>

	Private Function EnumWMIRegistryProcessing(ByVal objReg, ByVal strRegistryHive, ByVal strRegistryKey, ByVal strMethodToCall, _
													ByVal blnIs64BitMachine, ByVal blnIsWow6432Node)
	'*****************************************************************************************************************************************
	'*  Purpose:				Creates and execute registry enumeration processing
	'*  Arguments supplied:		Look up
	'*  Return Value:			Whatever the Method call returns (0 indicates success; non-zero indicates failure)
	'*  Called by:				ProcessWMIRegistryKeys, EnumWMIRegistryEntries, GetWMIRegistryEntry, GetWMIRegistryEntrySpecific
	'*  Calls:					LogThis, GetExpandedHiveInformation
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim lngRegistryHive, strExpandedRegistryHive, objNamedValueSet, objInputParams, objOutputParams, intErrNumber, strErrDescription
		Dim intCount, strError

		Call LogThis(vbTab & vbTab & "Enumerating (Exact) WMI registry for " & strRegistryHive & "\" & strRegistryKey, m_objLogAndTrace)
		Call GetExpandedHiveInformation(strRegistryHive, lngRegistryHive, strExpandedRegistryHive)
		If (blnIs64BitMachine) Then
			If (blnIsWow6432Node) Then
				'
				' Setup NamedValueSet needed for 32-bit machines
				'
				Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
				objNamedValueSet.Add "__ProviderArchitecture", 32
				objNamedValueSet.Add "__RequiredArchitecture", True
			Else
				'
				' Setup NamedValueSet needed for 64-bit machines
				'
				Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
				objNamedValueSet.Add "__ProviderArchitecture", 64
				objNamedValueSet.Add "__RequiredArchitecture", True
			End If
		Else
			'
			' Setup NamedValueSet needed for 32-bit machines
			'
			Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
			objNamedValueSet.Add "__ProviderArchitecture", 32
			objNamedValueSet.Add "__RequiredArchitecture", True
		End If
		
		Set objInputParams = objReg.Methods_(strMethodToCall).Inparameters
		objInputParams.Hdefkey = lngRegistryHive
		objInputParams.Ssubkeyname = strRegistryKey
		'
		' Execute the method and get the output
		'
		On Error Resume Next
		Set objOutputParams = objReg.ExecMethod_(strMethodToCall, objInputParams, , objNamedValueSet)
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0

		EnumWMIRegistryProcessing = -1
		If (intErrNumber = 0) Then
			If ((objOutputParams.ReturnValue = 0) And (IsArray(objOutputParams.sNames))) Then
				For intCount = 0 To UBound(objOutputParams.sNames)
					m_rsRegistryEnum.AddNew
					Call g_objFunctions.LoadRS(m_rsRegistryEnum, "Hive", strRegistryHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEnum, "Key", strRegistryKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEnum, "Subkey", objOutputParams.sNames(intCount), m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					If (UCase(strMethodToCall) = "ENUMVALUES") Then
						Call g_objFunctions.LoadRS(m_rsRegistryEnum, "Type", objOutputParams.Types(intCount), m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					End If
					m_rsRegistryEnum.Update
					Call LogThis(vbTab & vbTab & "Found registry key " & objOutputParams.sNames(intCount) & " in " & strRegistryKey, m_objLogAndTrace)
				Next
				If (m_rsRegistryEnum.RecordCount > 0) Then
					EnumWMIRegistryProcessing = 0
				End If
			Else
				'
				' An error occurred if ReturnValue <> 2 (2 = Entry Not Found; 3 = Key Not Found)
				'
				If ((objOutputParams.ReturnValue <> 0) And (objOutputParams.ReturnValue <> 2) And (objOutputParams.ReturnValue <> 3)) Then
					strError = "An error occurred during registry read for " & strRegistryKey & " in EnumWMIRegistryProcessing.  " & _
								"Error: " & intErrNumber & "(Description: " & strErrDescription & ")   ReturnValue: " & objOutputParams.ReturnValue
					Call LogThis(vbTab & vbTab & strError, m_objLogAndTraceErrors)
				End If
			End If
		Else
			strError = "Systems error " & intErrNumber & "(Description: " & strErrDescription & ") occurred during registry read for " & _
							strRegistryKey & " in EnumWMIRegistryProcessing"
			Call LogThis(vbTab & vbTab & strError, m_objLogAndTraceErrors)
		End If

	End Function

	Private Function ExactEntryRegistryProcessing(ByVal objReg, ByVal strRegistryHive, ByVal strRegistryKey, ByVal strRegValueToFind, _
													ByVal strMethodToCall, ByVal blnIs64BitMachine, ByVal blnIsWow6432Node, _
													ByRef intErrNumber, ByRef strErrDescription, ByRef valRegValue)
	'*****************************************************************************************************************************************
	'*  Purpose:				Creates and execute registry processing against 64-bit computers
	'*  Arguments supplied:		Look up
	'*  Return Value:			Whatever the Method call returns (0 indicates success; non-zero indicates failure)
	'*  Called by:				EnumWMIRegistryEntries, GetWMIRegistryEntry, GetWMIRegistryEntrySpecific
	'*  Calls:					LogThis, GetExpandedHiveInformation
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim lngRegistryHive, strExpandedRegistryHive, objNamedValueSet, objInputParams, objOutputParams, strError

		Call LogThis(vbTab & vbTab & "Enumerating (Exact) WMI registry entry for " & strRegistryHive & "\" & strRegistryKey, m_objLogAndTrace)
		Call GetExpandedHiveInformation(strRegistryHive, lngRegistryHive, strExpandedRegistryHive)
		If (blnIs64BitMachine) Then
			If (blnIsWow6432Node) Then
				'
				' Setup NamedValueSet needed for 32-bit machines
				'
				Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
				objNamedValueSet.Add "__ProviderArchitecture", 32
				objNamedValueSet.Add "__RequiredArchitecture", True
			Else
				'
				' Setup NamedValueSet needed for 64-bit machines
				'
				Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
				objNamedValueSet.Add "__ProviderArchitecture", 64
				objNamedValueSet.Add "__RequiredArchitecture", True
			End If
		Else
			'
			' Setup NamedValueSet needed for 32-bit machines
			'
			Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
			objNamedValueSet.Add "__ProviderArchitecture", 32
			objNamedValueSet.Add "__RequiredArchitecture", True
		End If

		Set objInputParams = objReg.Methods_(strMethodToCall).Inparameters
		objInputParams.Hdefkey = lngRegistryHive
		objInputParams.Ssubkeyname = strRegistryKey
		objInputParams.Svaluename = strRegValueToFind
		'
		' Execute the method and get the output
		'
		On Error Resume Next
		Set objOutputParams = objReg.ExecMethod_(strMethodToCall, objInputParams, , objNamedValueSet)
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0

		ExactEntryRegistryProcessing = -1
		If (intErrNumber = 0) Then
			If (objOutputParams.ReturnValue = 0) Then
				Select Case UCase(strMethodToCall)
					Case "GETSTRINGVALUE", "GETEXPANDEDSTRINGVALUE", "GETMULTISTRINGVALUE"
						valRegValue = objOutputParams.sValue
					Case "GETDWORDVALUE", "GETBINARYVALUE"
						valRegValue = objOutputParams.uValue
				End Select
				Call LogThis(vbTab & vbTab & "Found registry value " & strRegValueToFind & " in " & strRegistryKey, m_objLogAndTrace)
				ExactEntryRegistryProcessing = 0
			Else
				'
				' An error occurred if ReturnValue <> 2 (2 = Entry Not Found; 3 = Key Not Found)
				'
				If ((objOutputParams.ReturnValue <> 2) And (objOutputParams.ReturnValue <> 3)) Then
					strError = "An error occurred during registry read for value " & strRegValueToFind & " in key " & strRegistryKey & _
									" in ExactEntryRegistryProcessing.  " & "Error: " & intErrNumber & "(Description: " & strErrDescription & _
									")   ReturnValue: " & objOutputParams.ReturnValue & "  " & strMethodToCall
					Call LogThis(vbTab & vbTab & strError, m_objLogAndTraceErrors)
				End If
			End If
		Else
			strError = "Systems error " & intErrNumber & "(Description: " & strErrDescription & ") occurred during registry read for " & _
							strRegistryKey & " in ExactEntryRegistryProcessing"
			Call LogThis(vbTab & vbTab & strError, m_objLogAndTraceErrors)
		End If
		'
		' Cleanup
		'
		Set objNamedValueSet = Nothing
		Set objInputParams = Nothing
		Set objOutputParams = Nothing

	End Function

	Private Function EnumWMIRegistryEntries(ByVal objRemoteRegServer, ByVal strRegistryHive, ByVal strRegistryKey, _
												ByVal blnIs64BitMachine, ByVal blnIsWow6432Node)
	'*****************************************************************************************************************************************
	'*  Purpose:				Enumerate all entries under a particular registry key
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate completion
	'*  Called by:				EnumRegistryEntries, EnumRegistryEntriesByKey
	'*  Calls:					DeleteAllRecordsetRows, EnumWMIRegistryProcessing, ExactEntryRegistryProcessing
	'*  Requirements:			Registry Constants
	'*****************************************************************************************************************************************
		Dim strHive, strKey, strSubkey, intType, strType, intRetVal, intErrNumber, strErrDescription, valRegValue, strRegVal, intIndex

		Call g_objFunctions.DeleteAllRecordsetRows(m_rsRegistryEnum)
		Call EnumWMIRegistryProcessing(objRemoteRegServer, strRegistryHive, strRegistryKey, "EnumValues", blnIs64BitMachine, blnIsWow6432Node)

		If (m_rsRegistryEnum.RecordCount > 0) Then
			If (Not m_rsRegistryEnum.BOF) Then
				m_rsRegistryEnum.MoveFirst
			End If
			While Not m_rsRegistryEnum.EOF
				strHive = m_rsRegistryEnum("Hive")
				strKey = m_rsRegistryEnum("Key")
				'
				' strSubkey will hold the Value Name from the Entry enumeration
				'
				strSubkey = m_rsRegistryEnum("Subkey")
				intType = m_rsRegistryEnum("Type")
				'
				' Make sure the returned information (subkey) contains valid data.  If there is a (default) value that was
				' added with no data (normally would be (value not set)) then it is returned during enumeration processing
				' with an invalid Subkey.  Skip it if that is the case.
				'
				If ((Not IsNull(strSubkey)) And (Not IsEmpty(strSubkey)) And (strSubkey <> "") And (strSubkey <> " ")) Then
					Select Case intType
						Case m_REG_SZ
							strType = "REG_SZ"
							intRetVal = ExactEntryRegistryProcessing(objRemoteRegServer, strRegistryHive, strKey, strSubkey, _
																		"GetStringValue", blnIs64BitMachine, blnIsWow6432Node, _
																		intErrNumber, strErrDescription, valRegValue)
						Case m_REG_EXPAND_SZ
							strType = "REG_EXPAND_SZ"
							intRetVal = ExactEntryRegistryProcessing(objRemoteRegServer, strRegistryHive, strKey, strSubkey, _
																		"GetExpandedStringValue", blnIs64BitMachine, blnIsWow6432Node, _
																		intErrNumber, strErrDescription, valRegValue)
						Case m_REG_BINARY
							strType = "REG_BINARY"
							intRetVal = ExactEntryRegistryProcessing(objRemoteRegServer, strRegistryHive, strKey, strSubkey, _
																		"GetBinaryValue", blnIs64BitMachine, blnIsWow6432Node, _
																		intErrNumber, strErrDescription, valRegValue)
							'
							' Build a string that holds the Binary values
							'
							strRegVal = ""
							For intIndex = 0 to Ubound(valRegValue)
								strRegVal = strRegVal & Right("0" & Hex(valRegValue(intIndex)), 2)
							Next
							'
							' Strip any space from the end of the string
							'
							valRegValue = RTrim(strRegVal)
						Case m_REG_DWORD
							strType = "REG_DWORD"
							intRetVal = ExactEntryRegistryProcessing(objRemoteRegServer, strRegistryHive, strKey, strSubkey, _
																		"GetDWORDValue", blnIs64BitMachine, blnIsWow6432Node, _
																		intErrNumber, strErrDescription, valRegValue)
							strRegVal = ""
							If (IsArray(valRegValue)) Then
								For intIndex = 0 To UBound(valRegValue)
									strRegVal = strRegVal & valRegValue(intIndex) & ";"
								Next
								'
								' Strip the separator (;) from the end of the string
								'
								If (Len(strRegVal) > 0) Then
									valRegValue = Trim(Left(strRegVal, Len(strRegVal) - 1))
								Else
									valRegValue = ""
								End If
							End If
						Case m_REG_MULTI_SZ
							strType = "REG_MULTI_SZ"
							intRetVal = ExactEntryRegistryProcessing(objRemoteRegServer, strRegistryHive, strKey, strSubkey, _
																		"GetMultiStringValue", blnIs64BitMachine, blnIsWow6432Node, _
																		intErrNumber, strErrDescription, valRegValue)
							If (IsArray(valRegValue)) Then
								For intIndex = 0 To UBound(valRegValue)
									strRegVal = strRegVal & valRegValue(intIndex) & ";"
								Next
								'
								' Strip the separator (;) from the end of the string
								'
								If (Len(strRegVal) > 0) Then
									valRegValue = Trim(Left(strRegVal, Len(strRegVal) - 1))
								Else
									valRegValue = ""
								End If
							End If
						Case Else
							strType = "UNKNOWN"
							valRegValue = ""
					End Select
					m_rsRegistryEntries.AddNew
					Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Hive", strRegistryHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Key", strKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Entry", strSubkey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Type", intType, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEntries, "TypeText", strType, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Data", valRegValue, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					m_rsRegistryEntries.Update
					Call LogThis(vbTab & vbTab & "Found registry entry " & strSubkey & " in " & strKey & _
									" that contains the value: (" & valRegValue & ")", m_objLogAndTrace)
				End If
				m_rsRegistryEnum.MoveNext
			Wend
		End If

	End Function

	Private Function GetWMIRegistryDefaultValue(ByVal objRemoteRegServer, ByVal strRegistryHive, ByVal strRegistryKey, _
													ByVal blnIs64BitMachine, ByVal blnIsWow6432Node)
	'*****************************************************************************************************************************************
	'*  Purpose:				Gets the specified registry value
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate completion
	'*  Called by:				GetRegistryValue
	'*  Calls:					LogThis, ExactEntryRegistryProcessing
	'*  Requirements:			Registry Constants
	'*****************************************************************************************************************************************
		Dim intRetVal, intErrNumber, strErrDescription, valRegValue
		
		Call LogThis(vbTab & vbTab & "Getting WMI registry entry for (Default) in " & _
						strRegistryHive & "\" & strRegistryKey, m_objLogAndTrace)

		intRetVal = ExactEntryRegistryProcessing(objRemoteRegServer, strRegistryHive, strRegistryKey, Null, _
													"GetStringValue", blnIs64BitMachine, blnIsWow6432Node, intErrNumber, _
													strErrDescription, valRegValue)
		If ((intRetVal = 0) And (intErrNumber = 0)) Then
			If (IsNull(valRegValue)) Then
				valRegValue = "NULL"
			End If
			m_rsRegistryEntries.AddNew
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Hive", strRegistryHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Key", strRegistryKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Entry", "(Default)", m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Type", m_REG_SZ, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "TypeText", "REG_SZ", m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Data", valRegValue, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			m_rsRegistryEntries.Update
			Call LogThis(vbTab & vbTab & "Found registry entry (Default) in " & strRegistryKey & _
							" that contains the value: (" & valRegValue & ")", m_objLogAndTrace)
			Exit Function
		End If
		intRetVal = ExactEntryRegistryProcessing(objRemoteRegServer, strRegistryHive, strRegistryKey, Null, _
													"GetDWordValue", blnIs64BitMachine, blnIsWow6432Node, intErrNumber, _
													strErrDescription, valRegValue)
		If ((intRetVal = 0) And (intErrNumber = 0)) Then
			If (IsNull(valRegValue)) Then
				valRegValue = "NULL"
			End If
			m_rsRegistryEntries.AddNew
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Hive", strRegistryHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Key", strRegistryKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Entry", "(Default)", m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Type", m_REG_DWORD, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "TypeText", "REG_DWORD", m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Data", valRegValue, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			m_rsRegistryEntries.Update
			Call LogThis(vbTab & vbTab & "Found registry entry (Default) in " & strRegistryKey & _
							" that contains the value: (" & valRegValue & ")", m_objLogAndTrace)
			Exit Function
		End If
		intRetVal = ExactEntryRegistryProcessing(objRemoteRegServer, strRegistryHive, strRegistryKey, Null, _
													"GetExpandedStringValue", blnIs64BitMachine, blnIsWow6432Node, intErrNumber, _
													strErrDescription, valRegValue)
		If ((intRetVal = 0) And (intErrNumber = 0)) Then
			If (IsNull(valRegValue)) Then
				valRegValue = "NULL"
			End If
			m_rsRegistryEntries.AddNew
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Hive", strRegistryHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Key", strRegistryKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Entry", "(Default)", m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Type", m_REG_EXPAND_SZ, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "TypeText", "REG_EXPAND_SZ", m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Data", valRegValue, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			m_rsRegistryEntries.Update
			Call LogThis(vbTab & vbTab & "Found registry entry (Default) in " & strRegistryKey & _
							" that contains the value: (" & valRegValue & ")", m_objLogAndTrace)
			Exit Function
		End If
		intRetVal = ExactEntryRegistryProcessing(objRemoteRegServer, strRegistryHive, strRegistryKey, Null, _
													"GetBinaryValue", blnIs64BitMachine, blnIsWow6432Node, intErrNumber, _
													strErrDescription, valRegValue)
		If ((intRetVal = 0) And (intErrNumber = 0)) Then
			If (IsNull(valRegValue)) Then
				valRegValue = "NULL"
			End If
			m_rsRegistryEntries.AddNew
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Hive", strRegistryHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Key", strRegistryKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Entry", "(Default)", m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Type", m_REG_BINARY, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "TypeText", "REG_BINARY", m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Data", valRegValue, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			m_rsRegistryEntries.Update
			Call LogThis(vbTab & vbTab & "Found registry entry (Default) in " & strRegistryKey & _
							" that contains the value: (" & valRegValue & ")", m_objLogAndTrace)
			Exit Function
		End If
		intRetVal = ExactEntryRegistryProcessing(objRemoteRegServer, strRegistryHive, strRegistryKey, Null, _
													"GetMultiStringValue", blnIs64BitMachine, blnIsWow6432Node, intErrNumber, _
													strErrDescription, valRegValue)
		If ((intRetVal = 0) And (intErrNumber = 0)) Then
			If (IsNull(valRegValue)) Then
				valRegValue = "NULL"
			End If
			m_rsRegistryEntries.AddNew
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Hive", strRegistryHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Key", strRegistryKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Entry", "(Default)", m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Type", m_REG_MULTI_SZ, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "TypeText", "REG_MULTI_SZ", m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Data", valRegValue, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
			m_rsRegistryEntries.Update
			Call LogThis(vbTab & vbTab & "Found registry entry (Default) in " & strRegistryKey & _
							" that contains the value: (" & valRegValue & ")", m_objLogAndTrace)
			Exit Function
		End If
		'
		' If we made it here then we didn't find the default entry
		'
		Call LogThis(vbTab & vbTab & "WMI registry default value in " & strRegistryKey & " not found", m_objLogAndTrace)

	End Function

	Private Function SetWMIRegistryEntry(ByVal objReg, ByVal strRegistryHive, ByVal strRegistryKey, ByVal strValueName, ByVal valToSet, _
											ByVal strMethodToCall, ByVal blnIs64BitMachine, ByVal blnIsWow6432Node)
	'*****************************************************************************************************************************************
	'*  Purpose:				Sets the registry entry to the specified value (if the entry doesn't exist it will be created)
	'*  Arguments supplied:		Look up
	'*  Return Value:			Whatever the Method call returns (0 indicates success; non-zero indicates failure)
	'*  Called by:				SetRegistryEntry
	'*  Calls:					LogThis, GetExpandedHiveInformation
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim lngRegistryHive, strExpandedRegistryHive, objNamedValueSet, objInputParams, objOutputParams, intErrNumber, intCount, strError

		Call LogThis(vbTab & vbTab & "Set WMI registry for " & strRegistryHive & "\" & strRegistryKey & "\" & strValueName, m_objLogAndTrace)
		Call GetExpandedHiveInformation(strRegistryHive, lngRegistryHive, strExpandedRegistryHive)
		SetWMIRegistryEntry = -1
		If (blnIs64BitMachine) Then
			If (blnIsWow6432Node) Then
				'
				' Setup NamedValueSet needed for 32-bit machines
				'
				Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
				objNamedValueSet.Add "__ProviderArchitecture", 32
				objNamedValueSet.Add "__RequiredArchitecture", True
			Else
				'
				' Setup NamedValueSet needed for 64-bit machines
				'
				Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
				objNamedValueSet.Add "__ProviderArchitecture", 64
				objNamedValueSet.Add "__RequiredArchitecture", True
			End If
		Else
			'
			' Setup NamedValueSet needed for 32-bit machines
			'
			Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
			objNamedValueSet.Add "__ProviderArchitecture", 32
			objNamedValueSet.Add "__RequiredArchitecture", True
		End If
		
		Set objInputParams = objReg.Methods_(strMethodToCall).Inparameters
		objInputParams.hDefKey = lngRegistryHive
		objInputParams.sSubKeyName = strRegistryKey
		objInputParams.SValueName = strValueName
		Select Case UCase(strMethodToCall)
			Case "SETSTRINGVALUE", "SETEXPANDEDSTRINGVALUE", "SETMULTISTRINGVALUE"
				objInputParams.sValue = valToSet
			Case "SETQWORDVALUE", "SETDWORDVALUE", "SETBINARYVALUE"
				objInputParams.uValue = valToSet
		End Select
		'
		' Execute the method and get the output
		'
		On Error Resume Next
		Set objOutputParams = objReg.ExecMethod_(strMethodToCall, objInputParams, , objNamedValueSet)
		intErrNumber = Err.Number
		On Error GoTo 0

		If (intErrNumber = 0) Then
			SetWMIRegistryEntry = objOutputParams.ReturnValue
			If (objOutputParams.ReturnValue <> 0) Then
				'
				' An error occurred
				'
				strError = "An error occurred during registry add for " & strValueName & " in SetWMIRegistryEntry.  " & _
							"Error: " & intErrNumber & "   ReturnValue: " & objOutputParams.ReturnValue
				Call LogThis(vbTab & vbTab & strError, m_objLogAndTraceErrors)
			End If
		Else
			strError = "Systems error " & intErrNumber & " occurred during registry add for " & strValueName & " in SetWMIRegistryEntry"
			Call LogThis(vbTab & vbTab & strError, m_objLogAndTraceErrors)
		End If

	End Function

	Private Function DeleteWMIRegistryEntry(ByVal objReg, ByVal strRegistryHive, ByVal strRegistryKey, ByVal strValueName, _
												ByVal blnIs64BitMachine, ByVal blnIsWow6432Node)
	'*****************************************************************************************************************************************
	'*  Purpose:				Deletes the specified registry entry.
	'*  Arguments supplied:		Look up
	'*  Return Value:			Whatever the Method call returns (0 indicates success; non-zero indicates failure)
	'*  Called by:				SetRegistryEntry
	'*  Calls:					LogThis, GetExpandedHiveInformation
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim lngRegistryHive, strExpandedRegistryHive, objNamedValueSet, objInputParams, objOutputParams, intErrNumber, intCount, strError

		Call LogThis(vbTab & vbTab & "Delete WMI registry for " & strRegistryHive & "\" & strRegistryKey & "\" & strValueName, m_objLogAndTrace)
		Call GetExpandedHiveInformation(strRegistryHive, lngRegistryHive, strExpandedRegistryHive)
		DeleteWMIRegistryEntry = -1
		If (blnIs64BitMachine) Then
			If (blnIsWow6432Node) Then
				'
				' Setup NamedValueSet needed for 32-bit machines
				'
				Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
				objNamedValueSet.Add "__ProviderArchitecture", 32
				objNamedValueSet.Add "__RequiredArchitecture", True
			Else
				'
				' Setup NamedValueSet needed for 64-bit machines
				'
				Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
				objNamedValueSet.Add "__ProviderArchitecture", 64
				objNamedValueSet.Add "__RequiredArchitecture", True
			End If
		Else
			'
			' Setup NamedValueSet needed for 32-bit machines
			'
			Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
			objNamedValueSet.Add "__ProviderArchitecture", 32
			objNamedValueSet.Add "__RequiredArchitecture", True
		End If

		Set objInputParams = objReg.Methods_("DeleteValue").Inparameters
		objInputParams.hDefKey = lngRegistryHive
		objInputParams.sSubKeyName = strRegistryKey
		objInputParams.SValueName = strValueName
		'
		' Execute the method and get the output
		'
		On Error Resume Next
		Set objOutputParams = objReg.ExecMethod_("DeleteValue", objInputParams, , objNamedValueSet)
		intErrNumber = Err.Number
		On Error GoTo 0

		If (intErrNumber = 0) Then
			DeleteWMIRegistryEntry = objOutputParams.ReturnValue
			If (objOutputParams.ReturnValue <> 0) Then
				'
				' An error occurred
				'
				strError = "An error occurred during registry delete for " & strValueName & " in DeleteWMIRegistryEntry.  " & _
							"Error: " & intErrNumber & "   ReturnValue: " & objOutputParams.ReturnValue
				Call LogThis(vbTab & vbTab & strError, m_objLogAndTraceErrors)
			End If
		Else
			strError = "Systems error " & intErrNumber & " occurred during registry delete for " & strValueName & " in DeleteWMIRegistryEntry"
			Call LogThis(vbTab & vbTab & strError, m_objLogAndTraceErrors)
		End If

	End Function

	Private Function DeleteWMIRegistryKey(ByVal objReg, ByVal strRegistryHive, ByVal strRegistryKey, ByVal blnIs64BitMachine, _
												ByVal blnIsWow6432Node)
	'*****************************************************************************************************************************************
	'*  Purpose:				Deletes the specified registry entry.
	'*  Arguments supplied:		Look up
	'*  Return Value:			Whatever the Method call returns (0 indicates success; non-zero indicates failure)
	'*  Called by:				SetRegistryEntry
	'*  Calls:					LogThis, GetExpandedHiveInformation
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim lngRegistryHive, strExpandedRegistryHive, objNamedValueSet, objInputParams, objOutputParams, intErrNumber, intCount, strError

		Call LogThis(vbTab & vbTab & "Delete WMI registry for " & strRegistryHive & "\" & strRegistryKey, m_objLogAndTrace)
		Call GetExpandedHiveInformation(strRegistryHive, lngRegistryHive, strExpandedRegistryHive)
		DeleteWMIRegistryKey = -1
		If (blnIs64BitMachine) Then
			If (blnIsWow6432Node) Then
				'
				' Setup NamedValueSet needed for 32-bit machines
				'
				Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
				objNamedValueSet.Add "__ProviderArchitecture", 32
				objNamedValueSet.Add "__RequiredArchitecture", True
			Else
				'
				' Setup NamedValueSet needed for 64-bit machines
				'
				Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
				objNamedValueSet.Add "__ProviderArchitecture", 64
				objNamedValueSet.Add "__RequiredArchitecture", True
			End If
		Else
			'
			' Setup NamedValueSet needed for 32-bit machines
			'
			Set objNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
			objNamedValueSet.Add "__ProviderArchitecture", 32
			objNamedValueSet.Add "__RequiredArchitecture", True
		End If

		Set objInputParams = objReg.Methods_("DeleteKey").Inparameters
		objInputParams.hDefKey = lngRegistryHive
		objInputParams.sSubKeyName = strRegistryKey
		'
		' Execute the method and get the output
		'
		On Error Resume Next
		Set objOutputParams = objReg.ExecMethod_("DeleteKey", objInputParams, , objNamedValueSet)
		intErrNumber = Err.Number
		On Error GoTo 0

		If (intErrNumber = 0) Then
			DeleteWMIRegistryKey = objOutputParams.ReturnValue
			If (objOutputParams.ReturnValue <> 0) Then
				'
				' An error occurred
				'
				strError = "An error occurred during registry delete for " & strRegistryKey & " in DeleteWMIRegistryKey.  " & _
							"Error: " & intErrNumber & "   ReturnValue: " & objOutputParams.ReturnValue
				Call LogThis(vbTab & vbTab & strError, m_objLogAndTraceErrors)
			End If
		Else
			strError = "Systems error " & intErrNumber & " occurred during registry delete for " & strRegistryKey & " in DeleteWMIRegistryKey"
			Call LogThis(vbTab & vbTab & strError, m_objLogAndTraceErrors)
		End If

	End Function

'#endregion

'#region <Remote Registry Processing Functions>

	Private Function EnumRemoteRegistryKeys(ByVal strConnectedWithThis, ByVal strRegistryHive, ByVal strRegistryKey)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get the registry information requested using RemoteRegistry calls
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				GetRemoteRegistryKey, EnumRegistryKeys, EnumRegistrySubkeysByKey
	'*  Calls:					LogThis, GetExpandedHiveInformation
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim objShell, strFullKey, strKeyToSearch, objExec, strStandardOut, lngRegistryHive, strExpandedRegistryHive, strParseKey, arrLines
		Dim intCount, strLine, strSubkey

		Call LogThis(vbTab & vbTab & "Enumerating remote registry keys for " & strRegistryHive & "\" & strRegistryKey, m_objLogAndTrace)
		Set objShell = CreateObject("Wscript.Shell")
		'
		' Build the strings needed for REQ QUERY command
		'
		strFullKey = strRegistryHive & "\" & strRegistryKey
		strKeyToSearch = "\\" & strConnectedWithThis & "\" & strFullKey
		If (InStr(strKeyToSearch, " ")) Then
			strKeyToSearch = Chr(34) & strKeyToSearch & Chr(34)
		End If
		'
		' Standard Registry Examples:
		'	strRegistryHive = HKLM
		'	strKeyPath = SOFTWARE\Microsoft\Microsoft SQL Server\SESQLSERVER
		'	strFullKey = HKLM\SOFTWARE\Microsoft\Microsoft SQL Server\SESQLSERVER
		'	strKeyToSearch = \\52VDYDL3-SEA167.area52.afnoapps.usaf.mil\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server\SESQLSERVER
		'
		' Wow6432Node Registry Examples:
		'	strRegistryHive = HKLM
		'	strKeyPath = SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\SESQLSERVER
		'	strFullKey = HKLM\SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\SESQLSERVER
		'	strKeyToSearch = \\52VDYDL3-SEA167.area52.afnoapps.usaf.mil\HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\SESQLSERVER
		'
		Set objExec = objShell.Exec("cmd /c REG QUERY " & strKeyToSearch)
		strStandardOut = Trim(objExec.StdOut.ReadAll)
		'
		' Make sure the Reg Query was successful
		'
		' Errors:
		'	"ERROR: The system was unable to find the specified registry key or value."
		'	"ERROR: Invalid syntax."
		'
		If (InStr(1, strStandardOut, "ERROR: ", vbTextCompare) = 0) Then
			'
			' Remote registry processing is odd in that we pass in: HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
			' and we get back: HKey_Local_Machine\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
			'
			Call GetExpandedHiveInformation(strRegistryHive, lngRegistryHive, strExpandedRegistryHive)
			strParseKey = strExpandedRegistryHive & "\" & strRegistryKey
			'
			' Remove the original search string from the returned data
			'
			strStandardOut = Replace(strStandardOut, strParseKey & "\", "", 1, 2000, vbTextCompare)
			strStandardOut = Replace(strStandardOut, strParseKey, "", 1, 2000, vbTextCompare)

			If (InStr(strStandardOut, vbCrLf) > 0) Then
				arrLines = Split(strStandardOut, vbCrLf)
			Else
				arrLines = Array(strStandardOut)
			End If
			'
			' Process any keys we got back
			'
			For intCount = 0 To UBound(arrLines) - 1
				'
				' Process each line
				'
				strLine = Trim(arrLines(intCount))
				If ((InStr(1, strLine, "REG_SZ", vbTextCompare) = 0) And _
					(InStr(1, strLine, "REG_MULTI_SZ", vbTextCompare) = 0) And _
					(InStr(1, strLine, "REG_EXPAND_SZ", vbTextCompare) = 0) And _
					(InStr(1, strLine, "REG_DWORD", vbTextCompare) = 0) And _
					(InStr(1, strLine, "REG_BINARY", vbTextCompare) = 0) And _
					(InStr(1, strLine, "REG_NONE", vbTextCompare) = 0) And _
					(strLine <> "") And (strLine <> " ")) Then
					strSubkey = Trim(strLine)
					m_rsRegistryEnum.AddNew
					Call g_objFunctions.LoadRS(m_rsRegistryEnum, "Hive", strRegistryHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEnum, "Key", strRegistryKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEnum, "Subkey", strSubkey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					m_rsRegistryEnum.Update
					Call LogThis(vbTab & vbTab & "Found registry key " & strSubkey & " in " & strRegistryKey, m_objLogAndTrace)
				End If
			Next
		End If
		'
		' Cleanup
		'
		Set objShell = Nothing
		Set objExec = Nothing

	End Function

	Private Function EnumRemoteRegistryEntries(ByVal strConnectedWithThis, ByVal strRegistryHive, ByVal strRegistryKey)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get the registry information requested using RemoteRegistry calls
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				EnumRegistryEntries, EnumRegistryEntriesByKey
	'*  Calls:					LogThis, GetExpandedHiveInformation
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim objShell, lngRegistryHive, strExpandedRegistryHive, strToSearch, strKeyName, objExec, strStandardOut, strToCompare, arrLines
		Dim intCount, strLine, arrLine, strEntryType, strEntryName, arrTemp, intEntryType, strData

		Call LogThis(vbTab & vbTab & "Enumerating remote registry entries for " & strRegistryHive & "\" & strRegistryKey, m_objLogAndTrace)
		Set objShell = CreateObject("Wscript.Shell")
		Call GetExpandedHiveInformation(strRegistryHive, lngRegistryHive, strExpandedRegistryHive)
		'
		' Build the strings needed for REQ QUERY command
		'
		strToSearch = strRegistryHive & "\" & strRegistryKey
		strKeyName = "\\" & strConnectedWithThis & "\" & strToSearch
		If (InStr(strKeyName, " ")) Then
			strKeyName = Chr(34) & strKeyName & Chr(34)
		End If
		'
		' Standard Registry Examples:
		'	strRegistryHive = HKLM
		'	strKeyPath = SOFTWARE\Microsoft\Microsoft SQL Server\SESQLSERVER
		'	strFullKey = HKLM\SOFTWARE\Microsoft\Microsoft SQL Server\SESQLSERVER
		'	strKeyToSearch = \\52VDYDL3-SEA167.area52.afnoapps.usaf.mil\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server\SESQLSERVER
		'
		' Wow6432Node Registry Examples:
		'	strRegistryHive = HKLM
		'	strKeyPath = SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\SESQLSERVER
		'	strFullKey = HKLM\SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\SESQLSERVER
		'	strKeyToSearch = \\52VDYDL3-SEA167.area52.afnoapps.usaf.mil\HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\SESQLSERVER
		'
		Set objExec = objShell.Exec("cmd /c REG QUERY " & strKeyName & " /z")
		strStandardOut = Trim(objExec.StdOut.ReadAll)
		'
		' Make sure the Reg Query was successful
		'
		' Errors:
		'	"ERROR: The system was unable to find the specified registry key or value."
		'	"ERROR: Invalid syntax."
		'
		strToCompare = strExpandedRegistryHive & "\" & strRegistryKey
		If ((InStr(1, strStandardOut, "ERROR: ", vbTextCompare) = 0) And _
			(InStr(1, strStandardOut, strToCompare, vbTextCompare) > 0)) Then
			'
			' Load m_rsEntries with results
			'
			If (InStr(strStandardOut, vbCrLf) > 0) Then
				arrLines = Split(strStandardOut, vbCrLf)
			Else
				arrLines = Array(strStandardOut)
			End If
			'
			' Process any keys we got back
			'
			For intCount = 0 To UBound(arrLines) - 1
				'
				' Process each line
				'
				strLine = Trim(arrLines(intCount))
				If ((InStr(1, strLine, "REG_SZ", vbTextCompare) > 0) Or _
					(InStr(1, strLine, "REG_MULTI_SZ", vbTextCompare) > 0) Or _
					(InStr(1, strLine, "REG_EXPAND_SZ", vbTextCompare) > 0) Or _
					(InStr(1, strLine, "REG_DWORD", vbTextCompare) > 0) Or _
					(InStr(1, strLine, "REG_BINARY", vbTextCompare) > 0) Or _
					(InStr(1, strLine, "REG_NONE", vbTextCompare) > 0)) Then
					'
					' If we are here then one of the following is true:
					'	1) There are one or more entries under the specified key
					'	2) There is only the (Default) entry
					'
					' Examples:
					'	ErrorReportingDir    REG_SZ (1)    C:\Program Files\Microsoft SQL Server\100\Shared\ErrorDumps\
					'	EnableErrorReporting    REG_DWORD (4)    0x0
					'	(Default)    REG_SZ (1)
					'	(Default)    REG_SZ (1)    @SYS:DoesNotExist
					'
					' There will always be a value name (may contain a space), a registry type, type byte length, and value.
					'	If case #2 above is true then there will be no value present.  This must be taken into
					'	consideration during processing.
					'
					If (InStr(1, strLine, "REG_SZ", vbTextCompare) > 0) Then
						arrLine = Split(strLine, "REG_SZ", 2, vbTextCompare)
						strEntryType = "REG_SZ"
					ElseIf (InStr(1, strLine, "REG_MULTI_SZ", vbTextCompare) > 0) Then
						arrLine = Split(strLine, "REG_MULTI_SZ", 2, vbTextCompare)
						strEntryType = "REG_MULTI_SZ"
					ElseIf (InStr(1, strLine, "REG_EXPAND_SZ", vbTextCompare) > 0) Then
						arrLine = Split(strLine, "REG_EXPAND_SZ", 2, vbTextCompare)
						strEntryType = "REG_EXPAND_SZ"
					ElseIf (InStr(1, strLine, "REG_DWORD", vbTextCompare) > 0) Then
						arrLine = Split(strLine, "REG_DWORD", 2, vbTextCompare)
						strEntryType = "REG_DWORD"
					ElseIf (InStr(1, strLine, "REG_BINARY", vbTextCompare) > 0) Then
						arrLine = Split(strLine, "REG_BINARY", 2, vbTextCompare)
						strEntryType = "REG_BINARY"
					Else
						arrLine = Split(strLine, "REG_NONE", 2, vbTextCompare)
						strEntryType = "REG_NONE"
					End If
					strEntryName = Trim(arrLine(0))
					If (InStr(Trim(arrLine(1)), "  ")) Then
						'
						' Both TypeLength and Value exist (should happen almost all of the time)
						'
						arrTemp = Split(Trim(arrLine(1)), " ", 2)
						intEntryType = Replace(Replace(arrTemp(0), "(", ""), ")", "")
						strData = Trim(arrTemp(1))
						If (UCase(strEntryType) = "REG_DWORD") Then
							If (InStr(1, strData, "0x", vbTextCompare) > 0) Then
								strData = CLng(Replace(strData, "0x", "&H"))
							Else
								strData = CLng(strData)
							End If
						End If
					Else
						intEntryType = Replace(Replace(Trim(arrLine(1)), "(", ""), ")", "")
						strData = ""
					End If
					m_rsRegistryEntries.AddNew
					Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Hive", strRegistryHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Key", strRegistryKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Entry", strEntryName, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Type", intEntryType, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEntries, "TypeText", strEntryType, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsRegistryEntries, "Data", strData, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					m_rsRegistryEntries.Update
					Call LogThis(vbTab & vbTab & "Found registry entry " & strEntryName & " in " & strRegistryKey & _
									" that contains the value: (" & strData & ")", m_objLogAndTrace)
				End If
			Next
		End If
		'
		' Cleanup
		'
		Set objExec = Nothing
		Set objShell = Nothing

	End Function

	Private Function SetRemoteRegistryEntry(ByVal strConnectedWithThis, ByVal strRegistryHive, ByVal strRegistryKey, ByVal strValueName, _
												ByVal valToSet, ByVal strDataType)
	'*****************************************************************************************************************************************
	'*  Purpose:				Sets the registry entry to the specified value (if the entry doesn't exist it will be created)
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				EnumRegistryEntries, EnumRegistryEntriesByKey
	'*  Calls:					LogThis, GetExpandedHiveInformation
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim objShell, lngRegistryHive, strExpandedRegistryHive, strToSearch, strKeyName, objExec, strStandardOut, strToCompare, arrLines
		Dim intCount, strLine, arrLine, strEntryType, strEntryName, arrTemp, intEntryType, strData

		Call LogThis(vbTab & vbTab & "Set remote registry for " & strRegistryHive & "\" & strRegistryKey & "\" & strValueName, m_objLogAndTrace)
		Set objShell = CreateObject("Wscript.Shell")
		Call GetExpandedHiveInformation(strRegistryHive, lngRegistryHive, strExpandedRegistryHive)
		'
		' Build the strings needed for REQ ADD command
		'
		strToSearch = strRegistryHive & "\" & strRegistryKey
		strKeyName = "\\" & strConnectedWithThis & "\" & strToSearch
		If (InStr(strKeyName, " ")) Then
			strKeyName = Chr(34) & strKeyName & Chr(34)
		End If
		If (UCase(strDataType) = "REG_EXPAND_SZ") Then
			'
			' By replacing the percent sign with a carat and percent sign it will cause the actual value to
			' be stored in the Expanded String entry (instead of the converted value).
			'
			valToSet = Replace(valToSet, "%", "^%")
		End If
		Set objExec = objShell.Exec("cmd /c REG ADD " & strKeyName & " /v " & strValueName & " /t " & strDataType & " /d " & valToSet & " /f")
		strStandardOut = Trim(objExec.StdOut.ReadAll)
		'
		' Make sure the Reg Add was successful
		'
		' Errors:
		'	"ERROR: Invalid syntax."
		'	"ERROR: Invalid key name."
		'
		If (InStr(1, strStandardOut, "ERROR: ", vbTextCompare) = 0) Then
			SetRemoteRegistryEntry = 0
		Else
			SetRemoteRegistryEntry = -1
		End If
		'
		' Cleanup
		'
		Set objExec = Nothing
		Set objShell = Nothing

	End Function

	Private Function DeleteRemoteRegistryEntry(ByVal strConnectedWithThis, ByVal strRegistryHive, ByVal strRegistryKey, ByVal strValueName)
	'*****************************************************************************************************************************************
	'*  Purpose:				Sets the registry entry to the specified value (if the entry doesn't exist it will be created)
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				EnumRegistryEntries, EnumRegistryEntriesByKey
	'*  Calls:					LogThis, GetExpandedHiveInformation
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim objShell, lngRegistryHive, strExpandedRegistryHive, strToDelete, strKeyName, objExec, strStandardOut

		Call LogThis(vbTab & vbTab & "Delete remote registry value " & strRegistryHive & "\" & strRegistryKey & "\" & strValueName, m_objLogAndTrace)
		Set objShell = CreateObject("Wscript.Shell")
		Call GetExpandedHiveInformation(strRegistryHive, lngRegistryHive, strExpandedRegistryHive)
		'
		' Build the strings needed for REG DELETE command
		'
		strToDelete = strRegistryHive & "\" & strRegistryKey
		strKeyName = "\\" & strConnectedWithThis & "\" & strToDelete
		If (InStr(strKeyName, " ")) Then
			strKeyName = Chr(34) & strKeyName & Chr(34)
		End If
		Set objExec = objShell.Exec("cmd /c REG DELETE " & strKeyName & " /v " & strValueName & " /f")
		strStandardOut = Trim(objExec.StdOut.ReadAll)
		'
		' Make sure the Reg Add was successful
		'
		' Errors:
		'	"ERROR: Invalid syntax."
		'	"ERROR: Invalid key name."
		'
		If (InStr(1, strStandardOut, "ERROR: ", vbTextCompare) = 0) Then
			DeleteRemoteRegistryEntry = 0
		Else
			DeleteRemoteRegistryEntry = -1
		End If
		'
		' Cleanup
		'
		Set objExec = Nothing
		Set objShell = Nothing

	End Function

	Private Function DeleteRemoteRegistryKey(ByVal strConnectedWithThis, ByVal strRegistryHive, ByVal strRegistryKey)
	'*****************************************************************************************************************************************
	'*  Purpose:				Deletes the specified registry Key
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				EnumRegistryEntries, EnumRegistryEntriesByKey
	'*  Calls:					LogThis, GetExpandedHiveInformation
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim objShell, lngRegistryHive, strExpandedRegistryHive, strToDelete, strKeyName, objExec, strStandardOut

		Call LogThis(vbTab & vbTab & "Delete remote registry key " & strRegistryHive & "\" & strRegistryKey, m_objLogAndTrace)
		Set objShell = CreateObject("Wscript.Shell")
		Call GetExpandedHiveInformation(strRegistryHive, lngRegistryHive, strExpandedRegistryHive)
		'
		' Build the strings needed for REG DELETE command
		'
		strToDelete = strRegistryHive & "\" & strRegistryKey
		strKeyName = "\\" & strConnectedWithThis & "\" & strToDelete
		If (InStr(strKeyName, " ")) Then
			strKeyName = Chr(34) & strKeyName & Chr(34)
		End If
		Set objExec = objShell.Exec("cmd /c REG DELETE " & strKeyName & " /f")
		strStandardOut = Trim(objExec.StdOut.ReadAll)
		'
		' Make sure the Reg Add was successful
		'
		' Errors:
		'	"ERROR: Invalid syntax."
		'	"ERROR: Invalid key name."
		'
		If (InStr(1, strStandardOut, "ERROR: ", vbTextCompare) = 0) Then
			DeleteRemoteRegistryKey = 0
		Else
			DeleteRemoteRegistryKey = -1
		End If
		'
		' Cleanup
		'
		Set objExec = Nothing
		Set objShell = Nothing

	End Function

'#endregion

'#region <Shared Processing Functions>

	Private Function ParseRegistryEnumToKeys(ByVal blnIsWow6432Node, ByVal blnExactMatch, ByVal blnSpecificKey, ByVal strSearchKey)
	'*****************************************************************************************************************************************
	'*  Purpose:				Parse the m_rsRegistryEnum recordset
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate completion
	'*  Called by:				EnumRegistryKeys, GetRegistryKey
	'*  Calls:					None
	'*  Requirements:			Registry Constants
	'*****************************************************************************************************************************************
		Dim strHive, strKey, strSubkey, intType

		If (m_rsRegistryEnum.RecordCount > 0) Then
			If (Not m_rsRegistryEnum.BOF) Then
				m_rsRegistryEnum.MoveFirst
			End If
			While Not m_rsRegistryEnum.EOF
				strHive = m_rsRegistryEnum("Hive")
				strKey = m_rsRegistryEnum("Key")
				strSubkey = m_rsRegistryEnum("Subkey")
				intType = m_rsRegistryEnum("Type")
				If (blnSpecificKey) Then
					'
					' Just load a matching key
					'
					If (blnExactMatch) Then
						If (UCase(strSubkey) = UCase(strSearchKey)) Then
							m_rsKeys.AddNew
							Call g_objFunctions.LoadRS(m_rsKeys, "Hive", strHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsKeys, "Key", strKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsKeys, "Subkey", strSubkey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsKeys, "Wow6432Node", blnIsWow6432Node, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							m_rsKeys.Update
							Call LogThis(vbTab & vbTab & "Remote registry key " & strSearchKey & " exists", m_objLogAndTrace)
							Exit Function
						End If
					Else
						If (InStr(1, strSubkey, strSearchKey, vbTextCompare) > 0) Then
							m_rsKeys.AddNew
							Call g_objFunctions.LoadRS(m_rsKeys, "Hive", strHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsKeys, "Key", strKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsKeys, "Subkey", strSubkey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsKeys, "Wow6432Node", blnIsWow6432Node, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							m_rsKeys.Update
							Call LogThis(vbTab & vbTab & "Remote registry key " & strSubkey & " found", m_objLogAndTrace)
							Call LogThis(vbTab & vbTab & "Remote registry key " & strSearchKey & " exists", m_objLogAndTrace)
							Exit Function
						End If
					End If
				Else
					m_rsKeys.AddNew
					Call g_objFunctions.LoadRS(m_rsKeys, "Hive", strHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsKeys, "Key", strKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsKeys, "Subkey", strSubkey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsKeys, "Wow6432Node", blnIsWow6432Node, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					m_rsKeys.Update
				End If
				m_rsRegistryEnum.MoveNext
			Wend
		End If

	End Function

	Private Function ParseRegistryEntriesToEntries(ByVal blnIsWow6432Node, ByVal blnExactMatch, ByVal blnSpecificKey, ByVal strSearchEntry)
	'*****************************************************************************************************************************************
	'*  Purpose:				Parse the m_rsRegistryEntries recordset
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate completion
	'*  Called by:				EnumRegistryEntries, GetRegistryEntry, GetSpecificRegistryEntry, EnumRegistryEntriesByKey
	'*  Calls:					None
	'*  Requirements:			Registry Constants
	'*****************************************************************************************************************************************
		Dim strHive, strKey, strEntry, intType, strTypeText, valData

		If (m_rsRegistryEntries.RecordCount > 0) Then
			If (Not m_rsRegistryEntries.BOF) Then
				m_rsRegistryEntries.MoveFirst
			End If
			While Not m_rsRegistryEntries.EOF
				strHive = m_rsRegistryEntries("Hive")
				strKey = m_rsRegistryEntries("Key")
				strEntry = m_rsRegistryEntries("Entry")
				intType = m_rsRegistryEntries("Type")
				strTypeText = m_rsRegistryEntries("TypeText")
				valData = m_rsRegistryEntries("Data")
				If (blnSpecificKey) Then
					'
					' Just load a matching key
					'
					If (blnExactMatch) Then
						If (UCase(strEntry) = UCase(strSearchEntry)) Then
							m_rsEntries.AddNew
							Call g_objFunctions.LoadRS(m_rsEntries, "Hive", strHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsEntries, "Key", strKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsEntries, "Entry", strEntry, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsEntries, "Type", intType, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsEntries, "TypeText", strTypeText, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsEntries, "Wow6432Node", blnIsWow6432Node, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsEntries, "Data", valData, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							m_rsEntries.Update
							Call LogThis(vbTab & vbTab & "Remote registry entry " & strEntry & " exists", m_objLogAndTrace)
							Exit Function
						End If
					Else
						If (InStr(1, strEntry, strSearchEntry, vbTextCompare) > 0) Then
							m_rsEntries.AddNew
							Call g_objFunctions.LoadRS(m_rsEntries, "Hive", strHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsEntries, "Key", strKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsEntries, "Entry", strEntry, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsEntries, "Type", intType, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsEntries, "TypeText", strTypeText, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsEntries, "Wow6432Node", blnIsWow6432Node, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							Call g_objFunctions.LoadRS(m_rsEntries, "Data", valData, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
							m_rsEntries.Update
							Call LogThis(vbTab & vbTab & "Remote registry value " & strEntry & " found", m_objLogAndTrace)
							Call LogThis(vbTab & vbTab & "Remote registry value " & strSearchEntry & " exists", m_objLogAndTrace)
							Exit Function
						End If
					End If
				Else
					m_rsEntries.AddNew
					Call g_objFunctions.LoadRS(m_rsEntries, "Hive", strHive, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsEntries, "Key", strKey, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsEntries, "Entry", strEntry, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsEntries, "Type", intType, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsEntries, "TypeText", strTypeText, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsEntries, "Wow6432Node", blnIsWow6432Node, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					Call g_objFunctions.LoadRS(m_rsEntries, "Data", valData, m_objLogAndTraceLoadRS, m_objLogAndTraceErrors)
					m_rsEntries.Update
				End If
				m_rsRegistryEntries.MoveNext
			Wend
		End If

	End Function

'#endregion

'#region <Public Functions>

	Public Function RegKeySpecificEntryExists(ByVal objReg, ByVal strHive, ByVal strRegKey, ByVal strEntry)
	'*****************************************************************************************************************************************
	'*  Purpose:				Checks to see if a specific registry Entry exists
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success, non-zero to indicate failure (see WbemErrorEnum for errors)
	'*  Called by:				Mainline
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim arrEntryNames(), arrEntryTypes(), intRetVal, intErrNumber, intCount, strEntryName

		On Error Resume Next
		intRetVal = objReg.EnumValues(strHive, strRegKey, arrEntryNames, arrEntryTypes)
		intErrNumber = Err.Number
		On Error GoTo 0
		If (intErrNumber <> 0) Then
			For intCount = 0 To UBound(arrEntryNames) -1
				strEntryName = arrEntryNames(intCount)
				If (UCase(strEntryName) = UCase(strEntry)) Then
					RegKeySpecificEntryExists = True
					Exit Function
				End If
			Next
		End If
		RegKeySpecificEntryExists = intRetVal
		
	End Function

	Public Function RegKeySpecificSubkeyExists(ByVal objReg, ByVal strHive, ByVal strRegKey, ByVal strKey)
	'*****************************************************************************************************************************************
	'*  Purpose:				Checks to see if a specific registry Entry exists
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success, non-zero to indicate failure (see WbemErrorEnum for errors)
	'*  Called by:				Mainline
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim arrKeyNames(), intRetVal, intErrNumber, intCount, strKeyName

		On Error Resume Next
		intRetVal = objReg.EnumKeys(strHive, strRegKey, arrKeyNames)
		intErrNumber = Err.Number
		On Error GoTo 0
		If (intErrNumber <> 0) Then
			For intCount = 0 To UBound(arrKeyNames) -1
				strKeyName = arrKeyNames(intCount)
				If (UCase(strKeyName) = UCase(strKey)) Then
					RegKeySpecificSubkeyExists = True
					Exit Function
				End If
			Next
		End If
		RegKeySpecificSubkeyExists = intRetVal
		
	End Function

	Public Function RegKeyEntryExists(ByVal objReg, ByVal strHive, ByVal strRegKey)
	'*****************************************************************************************************************************************
	'*  Purpose:				Checks to see if any registry Entries exists
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success, non-zero to indicate failure (see WbemErrorEnum for errors)
	'*  Called by:				Mainline
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim arrEntryNames(), arrEntryTypes(), intRetVal, intErrNumber

		On Error Resume Next
		intRetVal = objReg.EnumValues(strHive, strRegKey, arrEntryNames, arrEntryTypes)
		intErrNumber = Err.Number
		On Error GoTo 0
		If (intErrNumber <> 0) Then
			RegKeyEntryExists = intRetVal
			Exit Function
		End If
		RegKeyEntryExists = intRetVal
		
	End Function

	Public Function RegKeySubkeyExists(ByVal objReg, ByVal strHive, ByVal strRegKey)
	'*****************************************************************************************************************************************
	'*  Purpose:				Checks to see if any registry Subkeys exists
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success, non-zero to indicate failure (see WbemErrorEnum for errors)
	'*  Called by:				Mainline
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim arrSubKeys(), intRetVal, intErrNumber

		On Error Resume Next
		intRetVal = objReg.EnumKeys(strHive, strRegKey, arrSubKeys)
		intErrNumber = Err.Number
		On Error GoTo 0
		If (intErrNumber <> 0) Then
			RegKeySubkeyExists = intRetVal
			Exit Function
		End If
		RegKeySubkeyExists = intRetVal
		
	End Function

	Public Function GetHiveInformation(ByVal strExpandedRegistryHive, ByRef lngRegistryHive, ByRef strRegistryHive)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get the associated hive information
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Main()
	'*  Calls:					None
	'*  Requirements:			Registry Constants
	'*****************************************************************************************************************************************

		Select Case strExpandedRegistryHive
			Case "HKEY_LOCAL_MACHINE"
				lngRegistryHive = m_HKEY_LOCAL_MACHINE
				strRegistryHive = "HKLM"
			Case "HKEY_CURRENT_USER"
				lngRegistryHive = m_HKEY_CURRENT_USER
				strRegistryHive = "HKCU"
			Case "HKEY_USERS"
				lngRegistryHive = m_HKEY_USERS
				strRegistryHive = "HKU"
			Case "HKEY_CLASSES_ROOT"
				lngRegistryHive = m_HKEY_CLASSES_ROOT
				strRegistryHive = "HKCR"
			Case "HKEY_CURRENT_CONFIG"
				lngRegistryHive = m_HKEY_CURRENT_CONFIG
				strRegistryHive = "HKCC"
		End Select

	End Function

	Public Function GetExpandedHiveInformation(ByVal strRegistryHive, ByRef lngRegistryHive, ByRef strExpandedRegistryHive)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get the associated hive information
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Main()
	'*  Calls:					None
	'*  Requirements:			Registry Constants
	'*****************************************************************************************************************************************

		Select Case strRegistryHive
			Case "HKLM"
				lngRegistryHive = m_HKEY_LOCAL_MACHINE
				strExpandedRegistryHive = "HKEY_LOCAL_MACHINE"
			Case "HKCU"
				lngRegistryHive = m_HKEY_CURRENT_USER
				strExpandedRegistryHive = "HKEY_CURRENT_USER"
			Case "HKU"
				lngRegistryHive = m_HKEY_USERS
				strExpandedRegistryHive = "HKEY_USERS"
			Case "HKCR"
				lngRegistryHive = m_HKEY_CLASSES_ROOT
				strExpandedRegistryHive = "HKEY_CLASSES_ROOT"
			Case "HKCC"
				lngRegistryHive = m_HKEY_CURRENT_CONFIG
				strExpandedRegistryHive = "HKEY_CURRENT_CONFIG"
		End Select

	End Function

	Public Function GetRegistryDataType(ByVal strOperation, ByRef strDataType)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get the associated hive information
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Main()
	'*  Calls:					None
	'*  Requirements:			Registry Constants
	'*****************************************************************************************************************************************

		Select Case UCase(strOperation)
			Case "STRING"
				strDataType = "REG_SZ"
			Case "MULTISTRING"
				strDataType = "REG_MULTI_SZ"
			Case "EXPANDEDSTRING"
				strDataType = "REG_EXPAND_SZ"
			Case "BINARY"
				strDataType = "REG_BINARY"
			Case "QWORD"
				strDataType = "REG_QWORD"
			Case "DWORD"
				strDataType = "REG_DWORD"
			Case Else
		End Select

	End Function

	Public Function GetMethodToCall(ByVal strOperation, ByRef strMethodToCall)
	'*****************************************************************************************************************************************
	'*  Purpose:				Gets the appropriate Add Method based on operation to be performed.
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Main()
	'*  Calls:					None
	'*  Requirements:			Registry Constants
	'*****************************************************************************************************************************************

		Select Case UCase(strOperation)
			Case "STRING"
				strMethodToCall = "SetStringValue"
			Case "MULTISTRING"
				strMethodToCall = "SetMultiStringValue"
			Case "EXPANDEDSTRING"
				strMethodToCall = "SetExpandedStringValue"
			Case "BINARY"
				strMethodToCall = "SetBinaryValue"
			Case "QWORD"
				strMethodToCall = "SetQWORDValue"
			Case "DWORD"
				strMethodToCall = "SetDWORDValue"
			Case Else
		End Select

	End Function

	Public Function EnumRegistryKeys(ByVal objRemoteRegServer, ByVal strConnectedWithThis, ByVal blnWMIRegGoodToGo, _
											ByVal blnRemoteRegGoodToGo, ByVal strRegistryHive, ByVal strRegistryKey, _
											ByVal blnIs64BitMachine, ByVal strOSVersion, ByRef blnKeyOrValueExists, ByRef rsKeys, _
											ByRef objLogAndTrace, ByRef objLogAndTraceErrors, ByRef objLogAndTraceLoadRS)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get all registry keys using WMI Registry or RemoteRegistry processing
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					LogThis, DeleteAllRecordsetRows, RequireTwoRegistryReads, EnumWMIRegistryProcessing, EnumRemoteRegistryKeys
	'*							ParseRegistryEnumToKeys, CreateDuplicateRecordset
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strKeyPath, blnTwoReadsRequired, blnIsWow6432Node, blnExactMatch, blnSpecificKey, strSearchKey

		If (IsObject(objLogAndTrace)) Then
			Set m_objLogAndTrace = objLogAndTrace
		End If
		If (IsObject(objLogAndTraceErrors)) Then
			Set m_objLogAndTraceErrors = objLogAndTraceErrors
		End If
		If (IsObject(objLogAndTraceLoadRS)) Then
			Set m_objLogAndTraceLoadRS = objLogAndTraceLoadRS
		End If

		Call LogThis(vbTab & "Processing EnumRegistryKeys started", m_objLogAndTrace)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsKeys)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsRegistryEnum)

		If (strRegistryKey = "") Then
			strKeyPath = ""
			blnTwoReadsRequired = False
		Else
			'
			' Strip off the '\' at the end of the string if present
			'
			If (Right(strRegistryKey, 1) = "\") Then
				strKeyPath = Left(strRegistryKey, Len(strRegistryKey) - 1)
			Else
				strKeyPath = strRegistryKey
			End If
			blnTwoReadsRequired = RequireTwoRegistryReads(strRegistryHive & "\" & strKeyPath, strOSVersion)
		End If
		blnIsWow6432Node = False
		blnExactMatch = False
		blnSpecificKey = False
		strSearchKey = ""
		'
		' The following functions load the m_rsRegistryEnum recordset with registry Key information
		'
		If (blnWMIRegGoodToGo) Then
			Call EnumWMIRegistryProcessing(objRemoteRegServer, strRegistryHive, strKeyPath, "EnumKey", blnIs64BitMachine, blnIsWow6432Node)
		Else
			Call EnumRemoteRegistryKeys(strConnectedWithThis, strRegistryHive, strKeyPath)
		End If
		'
		' Parse the m_rsRegistryEnum recordset for Keys that match
		'
		Call ParseRegistryEnumToKeys(blnIsWow6432Node, blnExactMatch, blnSpecificKey, strSearchKey)
		
		If (blnTwoReadsRequired) Then
			Call g_objFunctions.DeleteAllRecordsetRows(m_rsRegistryEnum)
			'
			' Enumerate the Wow6432Node section of the Registry
			'
			blnIsWow6432Node = True
			'
			' The following functions load the m_rsRegistryEnum recordset with registry Key information
			'
			If (blnWMIRegGoodToGo) Then
				Call EnumWMIRegistryProcessing(objRemoteRegServer, strRegistryHive, strKeyPath, "EnumKey", blnIs64BitMachine, blnIsWow6432Node)
			Else
				strKeyPath = Replace(strKeyPath, "SOFTWARE\", "SOFTWARE\Wow6432Node\", 1, 2, vbTextCompare)
				Call EnumRemoteRegistryKeys(strConnectedWithThis, strRegistryHive, strKeyPath)
			End If
			'
			' Parse the m_rsRegistryEnum recordset for values that match
			'
			Call ParseRegistryEnumToKeys(blnIsWow6432Node, blnExactMatch, blnSpecificKey, strSearchKey)
		End If
		If (m_rsKeys.RecordCount > 0) Then
			blnKeyOrValueExists = True
		Else
			blnKeyOrValueExists = False
		End If
		Call g_objFunctions.CreateDuplicateRecordset(m_rsKeys, rsKeys)

	End Function

	Public Function GetRegistryKey(ByVal objRemoteRegServer, ByVal strConnectedWithThis, ByVal blnWMIRegGoodToGo, _
										ByVal blnRemoteRegGoodToGo, ByVal strRegistryHive, ByVal strRegistryKey, ByVal strSearchKey, _
										ByVal blnExactMatch, ByVal blnIs64BitMachine, ByVal strOSVersion, ByRef rsKeys, _
										ByRef blnKeyOrValueExists, ByRef objLogAndTrace, ByRef objLogAndTraceErrors, _
										ByRef objLogAndTraceLoadRS)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get the registry key using WMI Registry or RemoteRegistry processing
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				None
	'*  Calls:					LogThis, DeleteAllRecordsetRows, RequireTwoRegistryReads, EnumWMIRegistryProcessing, EnumRemoteRegistryKeys
	'*							ParseRegistryEnumToKeys, CreateDuplicateRecordset
	'*	Requirements:			Registry Constants
	'*****************************************************************************************************************************************
		Dim strKeyPath, blnTwoReadsRequired, blnIsWow6432Node, blnSpecificKey

		If (IsObject(objLogAndTrace)) Then
			Set m_objLogAndTrace = objLogAndTrace
		End If
		If (IsObject(objLogAndTraceErrors)) Then
			Set m_objLogAndTraceErrors = objLogAndTraceErrors
		End If
		If (IsObject(objLogAndTraceLoadRS)) Then
			Set m_objLogAndTraceLoadRS = objLogAndTraceLoadRS
		End If

		Call LogThis(vbTab & "Processing GetRegistryKey started", m_objLogAndTrace)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsKeys)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsRegistryEnum)
		blnKeyOrValueExists = False

		If (strRegistryKey = "") Then
			strKeyPath = ""
			blnTwoReadsRequired = False
		Else
			'
			' Strip off the '\' at the end of the string if present
			'
			If (Right(strRegistryKey, 1) = "\") Then
				strKeyPath = Left(strRegistryKey, Len(strRegistryKey) - 1)
			Else
				strKeyPath = strRegistryKey
			End If
			blnTwoReadsRequired = RequireTwoRegistryReads(strRegistryHive & "\" & strKeyPath, strOSVersion)
		End If
		blnIsWow6432Node = False
		blnSpecificKey = True
		'
		' The following functions load the m_rsRegistryEnum recordset with registry Key information
		'
		If (blnWMIRegGoodToGo) Then
			Call EnumWMIRegistryProcessing(objRemoteRegServer, strRegistryHive, strKeyPath, "EnumKey", blnIs64BitMachine, blnIsWow6432Node)
		Else
			Call EnumRemoteRegistryKeys(strConnectedWithThis, strRegistryHive, strKeyPath)
		End If
		'
		' Parse the m_rsRegistryEnum recordset for Keys that match
		'
		Call ParseRegistryEnumToKeys(blnIsWow6432Node, blnExactMatch, blnSpecificKey, strSearchKey)
		If (m_rsKeys.RecordCount > 0) Then
			'
			' Found it
			'
			blnKeyOrValueExists = True
			Call g_objFunctions.CreateDuplicateRecordset(m_rsKeys, rsKeys)
			Exit Function
		End If
		'
		' Since we are looking for a specific Key, only continue processing if it hasn't been found yet
		'
		If (blnTwoReadsRequired) Then
			Call g_objFunctions.DeleteAllRecordsetRows(m_rsRegistryEnum)
			'
			' Enumerate the Wow6432Node section of the Registry
			'
			blnIsWow6432Node = True
			'
			' The following functions load the m_rsRegistryEnum recordset with registry Key information
			'
			If (blnWMIRegGoodToGo) Then
				Call EnumWMIRegistryProcessing(objRemoteRegServer, strRegistryHive, strKeyPath, "EnumKey", blnIs64BitMachine, blnIsWow6432Node)
			Else
				strKeyPath = Replace(strKeyPath, "SOFTWARE\", "SOFTWARE\Wow6432Node\", 1, 2, vbTextCompare)
				Call EnumRemoteRegistryKeys(strConnectedWithThis, strRegistryHive, strKeyPath)
			End If
			'
			' Parse the m_rsRegistryEnum recordset for values that match
			'
			Call ParseRegistryEnumToKeys(blnIsWow6432Node, blnExactMatch, blnSpecificKey, strSearchKey)
			If (m_rsKeys.RecordCount > 0) Then
				'
				' Found it
				'
				blnKeyOrValueExists = True
				Call g_objFunctions.CreateDuplicateRecordset(m_rsKeys, rsKeys)
				Exit Function
			End If
		End If

	End Function

	Public Function EnumRegistryEntries(ByVal objRemoteRegServer, ByVal strConnectedWithThis, ByVal blnWMIRegGoodToGo, _
											ByVal blnRemoteRegGoodToGo, ByVal strRegistryHive, ByVal strRegistryKey, _
											ByVal blnIs64BitMachine, ByVal strOSVersion, ByRef blnKeyOrValueExists, ByRef rsEntries, _
											ByRef objLogAndTrace, ByRef objLogAndTraceErrors, ByRef objLogAndTraceLoadRS)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get all registry entries using WMI Registry or RemoteRegistry processing
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					LogThis, DeleteAllRecordsetRows, RequireTwoRegistryReads, EnumWMIRegistryEntries, EnumRemoteRegistryEntries
	'*							ParseRegistryEntriesToEntries, CreateDuplicateRecordset
	'*	Requirements:			Registry Constants
	'*****************************************************************************************************************************************
		Dim strKeyPath, blnTwoReadsRequired, blnIsWow6432Node, blnExactMatch, blnSpecificEntry, strSearchEntry

		If (IsObject(objLogAndTrace)) Then
			Set m_objLogAndTrace = objLogAndTrace
		End If
		If (IsObject(objLogAndTraceErrors)) Then
			Set m_objLogAndTraceErrors = objLogAndTraceErrors
		End If
		If (IsObject(objLogAndTraceLoadRS)) Then
			Set m_objLogAndTraceLoadRS = objLogAndTraceLoadRS
		End If

		Call LogThis(vbTab & "Processing EnumRegistryEntries started", m_objLogAndTrace)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsEntries)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsRegistryEntries)

		If (strRegistryKey = "") Then
			strKeyPath = ""
			blnTwoReadsRequired = False
		Else
			'
			' Strip off the '\' at the end of the string if present
			'
			If (Right(strRegistryKey, 1) = "\") Then
				strKeyPath = Left(strRegistryKey, Len(strRegistryKey) - 1)
			Else
				strKeyPath = strRegistryKey
			End If
			blnTwoReadsRequired = RequireTwoRegistryReads(strRegistryHive & "\" & strKeyPath, strOSVersion)
		End If
		blnIsWow6432Node = False
		blnExactMatch = False
		blnSpecificEntry = False
		strSearchEntry = ""
		'
		' The following functions load the m_rsRegistryEntries recordset with registry Entry information
		'
		If (blnWMIRegGoodToGo) Then
			Call EnumWMIRegistryEntries(objRemoteRegServer, strRegistryHive, strKeyPath, blnIs64BitMachine, blnIsWow6432Node)
		ElseIf (blnRemoteRegGoodToGo) Then
			Call EnumRemoteRegistryEntries(strConnectedWithThis, strRegistryHive, strKeyPath)
		End If
		'
		' Parse the m_rsRegistryEntries recordset for entries that match
		'
		Call ParseRegistryEntriesToEntries(blnIsWow6432Node, blnExactMatch, blnSpecificEntry, strSearchEntry)

		If (blnTwoReadsRequired) Then
			Call g_objFunctions.DeleteAllRecordsetRows(m_rsRegistryEntries)
			'
			' Enumerate the Wow6432Node section of the Registry
			'
			blnIsWow6432Node = True
			'
			' The following functions load the m_rsRegistryEntries recordset with registry Entry information
			'
			If (blnWMIRegGoodToGo) Then
				Call EnumWMIRegistryEntries(objRemoteRegServer, strRegistryHive, strKeyPath, blnIs64BitMachine, blnIsWow6432Node)
			ElseIf (blnRemoteRegGoodToGo) Then
				strKeyPath = Replace(strKeyPath, "SOFTWARE\", "SOFTWARE\Wow6432Node\", 1, 2, vbTextCompare)
				Call EnumRemoteRegistryEntries(strConnectedWithThis, strRegistryHive, strKeyPath)
			End If
			'
			' Parse the m_rsRegistryEntries recordset for entries that match
			'
			Call ParseRegistryEntriesToEntries(blnIsWow6432Node, blnExactMatch, blnSpecificEntry, strSearchEntry)
		End If
		If (m_rsEntries.RecordCount > 0) Then
			blnKeyOrValueExists = True
		Else
			blnKeyOrValueExists = False
		End If
		Call g_objFunctions.CreateDuplicateRecordset(m_rsEntries, rsEntries)

	End Function

	Public Function GetRegistryEntry(ByVal objRemoteRegServer, ByVal strConnectedWithThis, ByVal blnWMIRegGoodToGo, _
										ByVal blnRemoteRegGoodToGo, ByVal strRegistryHive, ByVal strRegistryKey, ByVal strSearchEntry, _
										ByVal blnIs64BitMachine, ByVal strOSVersion, ByRef valRegValue, ByRef blnKeyOrValueExists, _
										ByRef strKeyType, ByRef objLogAndTrace, ByRef objLogAndTraceErrors, ByRef objLogAndTraceLoadRS)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get the registry value using WMI Registry or RemoteRegistry processing
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					LogThis, DeleteAllRecordsetRows, RequireTwoRegistryReads, EnumWMIRegistryEntries, EnumRemoteRegistryEntries
	'*							ParseRegistryEntriesToEntries
	'*	Requirements:			Registry Constants
	'*****************************************************************************************************************************************
		Dim strKeyPath, blnTwoReadsRequired, blnIsWow6432Node, blnSpecificEntry, blnExactMatch

		If (IsObject(objLogAndTrace)) Then
			Set m_objLogAndTrace = objLogAndTrace
		End If
		If (IsObject(objLogAndTraceErrors)) Then
			Set m_objLogAndTraceErrors = objLogAndTraceErrors
		End If
		If (IsObject(objLogAndTraceLoadRS)) Then
			Set m_objLogAndTraceLoadRS = objLogAndTraceLoadRS
		End If

		Call LogThis(vbTab & "Processing GetRegistryEntry started", m_objLogAndTrace)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsEntries)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsRegistryEntries)
		blnKeyOrValueExists = False
		valRegValue = ""
		strKeyType = ""

		If (strRegistryKey = "") Then
			strKeyPath = ""
			blnTwoReadsRequired = False
		Else
			'
			' Strip off the '\' at the end of the string if present
			'
			If (Right(strRegistryKey, 1) = "\") Then
				strKeyPath = Left(strRegistryKey, Len(strRegistryKey) - 1)
			Else
				strKeyPath = strRegistryKey
			End If
			blnTwoReadsRequired = RequireTwoRegistryReads(strRegistryHive & "\" & strKeyPath, strOSVersion)
		End If
		blnIsWow6432Node = False
		blnSpecificEntry = True
		blnExactMatch = True
		'
		' The following functions load the m_rsRegistryEntries recordset with registry Entry information
		'
		If (IsNull(strSearchEntry)) Then
			If (blnWMIRegGoodToGo) Then
				Call GetWMIRegistryDefaultValue(objRemoteRegServer, strRegistryHive, strKeyPath, blnIs64BitMachine, blnIsWow6432Node)
				blnSpecificEntry = False
			ElseIf (blnRemoteRegGoodToGo) Then
				Call EnumRemoteRegistryEntries(strConnectedWithThis, strRegistryHive, strKeyPath)
			End If
		Else
			If (blnWMIRegGoodToGo) Then
				Call EnumWMIRegistryEntries(objRemoteRegServer, strRegistryHive, strKeyPath, blnIs64BitMachine, blnIsWow6432Node)
			ElseIf (blnRemoteRegGoodToGo) Then
				Call EnumRemoteRegistryEntries(strConnectedWithThis, strRegistryHive, strKeyPath)
			End If
		End If
		'
		' Parse the m_rsRegistryEntries recordset for entries that match
		'
		Call ParseRegistryEntriesToEntries(blnIsWow6432Node, blnExactMatch, blnSpecificEntry, strSearchEntry)
		If (m_rsEntries.RecordCount > 0) Then
			'
			' Found it
			'
			m_rsEntries.MoveFirst
			blnKeyOrValueExists = True
			valRegValue = m_rsEntries("Data")
			strKeyType = m_rsEntries("TypeText")
			Exit Function
		End If
		'
		' Since we are looking for a specific Entry, only continue processing if it hasn't been found yet
		'
		If (blnTwoReadsRequired) Then
			Call g_objFunctions.DeleteAllRecordsetRows(m_rsRegistryEntries)
			'
			' Enumerate the Wow6432Node section of the Registry
			'
			blnIsWow6432Node = True
			'
			' The following functions load the m_rsRegistryEntries recordset with registry Entry information
			'
			If (IsNull(strSearchEntry)) Then
				If (blnWMIRegGoodToGo) Then
					Call GetWMIRegistryDefaultValue(objRemoteRegServer, strRegistryHive, strKeyPath, blnIs64BitMachine, blnIsWow6432Node)
					blnSpecificEntry = False
				ElseIf (blnRemoteRegGoodToGo) Then
					strKeyPath = Replace(strKeyPath, "SOFTWARE\", "SOFTWARE\Wow6432Node\", 1, 2, vbTextCompare)
					Call EnumRemoteRegistryEntries(strConnectedWithThis, strRegistryHive, strKeyPath)
				End If
			Else
				If (blnWMIRegGoodToGo) Then
					Call EnumWMIRegistryEntries(objRemoteRegServer, strRegistryHive, strKeyPath, blnIs64BitMachine, blnIsWow6432Node)
				ElseIf (blnRemoteRegGoodToGo) Then
					strKeyPath = Replace(strKeyPath, "SOFTWARE\", "SOFTWARE\Wow6432Node\", 1, 2, vbTextCompare)
					Call EnumRemoteRegistryEntries(strConnectedWithThis, strRegistryHive, strKeyPath)
				End If
			End If
			'
			' Parse the m_rsRegistryEntries recordset for entries that match
			'
			Call ParseRegistryEntriesToEntries(blnIsWow6432Node, blnExactMatch, blnSpecificEntry, strSearchEntry)
			If (m_rsEntries.RecordCount > 0) Then
				'
				' Found it
				'
				m_rsEntries.MoveFirst
				blnKeyOrValueExists = True
				valRegValue = m_rsEntries("Data")
				strKeyType = m_rsEntries("TypeText")
				Exit Function
			End If
		End If

	End Function

	Public Function GetSpecificRegistryEntry(ByVal objRemoteRegServer, ByVal strConnectedWithThis, ByVal blnWMIRegGoodToGo, _
												ByVal blnRemoteRegGoodToGo, ByVal strRegistryHive, ByVal strRegistryKey, _
												ByVal strSearchEntry, ByVal blnIs64BitMachine, ByVal blnIsWow6432Node, ByRef valRegValue, _
												ByRef blnKeyOrValueExists, ByRef strKeyType, ByRef objLogAndTrace, _
												ByRef objLogAndTraceErrors, ByRef objLogAndTraceLoadRS)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get the registry value using WMI Registry or RemoteRegistry processing
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					LogThis, DeleteAllRecordsetRows, EnumWMIRegistryEntries, EnumRemoteRegistryEntries
	'*							ParseRegistryEntriesToEntries
	'*	Requirements:			Registry Constants
	'*****************************************************************************************************************************************
		Dim strKeyPath, blnSpecificEntry, blnExactMatch

		If (IsObject(objLogAndTrace)) Then
			Set m_objLogAndTrace = objLogAndTrace
		End If
		If (IsObject(objLogAndTraceErrors)) Then
			Set m_objLogAndTraceErrors = objLogAndTraceErrors
		End If
		If (IsObject(objLogAndTraceLoadRS)) Then
			Set m_objLogAndTraceLoadRS = objLogAndTraceLoadRS
		End If

		Call LogThis(vbTab & "Processing GetSpecificRegistryEntry started", m_objLogAndTrace)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsEntries)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsRegistryEntries)
		blnKeyOrValueExists = False

		If (strRegistryKey = "") Then
			strKeyPath = ""
		Else
			'
			' Strip off the '\' at the end of the string if present
			'
			If (Right(strRegistryKey, 1) = "\") Then
				strKeyPath = Left(strRegistryKey, Len(strRegistryKey) - 1)
			Else
				strKeyPath = strRegistryKey
			End If
		End If
		blnSpecificEntry = True
		blnExactMatch = True
		'
		' The following functions load the m_rsRegistryEntries recordset with registry Entry information
		'
		If (blnWMIRegGoodToGo) Then
			Call EnumWMIRegistryEntries(objRemoteRegServer, strRegistryHive, strKeyPath, blnIs64BitMachine, blnIsWow6432Node)
		ElseIf (blnRemoteRegGoodToGo) Then
			Call EnumRemoteRegistryEntries(strConnectedWithThis, strRegistryHive, strKeyPath)
		End If
		'
		' Parse the m_rsRegistryEntries recordset for entries that match
		'
		Call ParseRegistryEntriesToEntries(blnIsWow6432Node, blnExactMatch, blnSpecificEntry, strSearchEntry)
		If (m_rsEntries.RecordCount > 0) Then
			'
			' Found it
			'
			m_rsEntries.MoveFirst
			blnKeyOrValueExists = True
			valRegValue = m_rsEntries("Data")
			strKeyType = m_rsEntries("TypeText")
			Exit Function
		End If

	End Function

	Public Function EnumRegistryEntriesByKey(ByVal objRemoteRegServer, ByVal strConnectedWithThis, ByVal blnWMIRegGoodToGo, _
												ByVal blnRemoteRegGoodToGo, ByVal strRegistryHive, ByVal strRegistryKey, _
												ByVal blnIs64BitMachine, ByVal blnIsWow6432Node, ByVal strOSVersion, _
												ByRef blnKeyOrValueExists, ByRef rsEntries, ByRef objLogAndTrace, _
												ByRef objLogAndTraceErrors, ByRef objLogAndTraceLoadRS)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get all registry entries using WMI Registry or RemoteRegistry processing
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					LogThis, DeleteAllRecordsetRows, EnumWMIRegistryEntries, EnumRemoteRegistryEntries
	'*							ParseRegistryEntriesToEntries, CreateDuplicateRecordset
	'*	Requirements:			Registry Constants
	'*****************************************************************************************************************************************
		Dim strKeyPath, blnExactMatch, blnSpecificEntry, strSearchEntry

		If (IsObject(objLogAndTrace)) Then
			Set m_objLogAndTrace = objLogAndTrace
		End If
		If (IsObject(objLogAndTraceErrors)) Then
			Set m_objLogAndTraceErrors = objLogAndTraceErrors
		End If
		If (IsObject(objLogAndTraceLoadRS)) Then
			Set m_objLogAndTraceLoadRS = objLogAndTraceLoadRS
		End If

		Call LogThis(vbTab & "Processing EnumRegistryEntriesByKey started", m_objLogAndTrace)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsEntries)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsRegistryEntries)

		If (strRegistryKey = "") Then
			strKeyPath = ""
		Else
			'
			' Strip off the '\' at the end of the string if present
			'
			If (Right(strRegistryKey, 1) = "\") Then
				strKeyPath = Left(strRegistryKey, Len(strRegistryKey) - 1)
			Else
				strKeyPath = strRegistryKey
			End If
		End If
		blnExactMatch = False
		blnSpecificEntry = False
		strSearchEntry = ""
		'
		' The following functions load the m_rsRegistryEntries recordset with registry Entry information
		'
		If (blnWMIRegGoodToGo) Then
			Call EnumWMIRegistryEntries(objRemoteRegServer, strRegistryHive, strKeyPath, blnIs64BitMachine, blnIsWow6432Node)
		ElseIf (blnRemoteRegGoodToGo) Then
			Call EnumRemoteRegistryEntries(strConnectedWithThis, strRegistryHive, strKeyPath)
		End If
		'
		' Parse the m_rsRegistryEntries recordset for entries that match
		'
		Call ParseRegistryEntriesToEntries(blnIsWow6432Node, blnExactMatch, blnSpecificEntry, strSearchEntry)
		If (m_rsEntries.RecordCount > 0) Then
			blnKeyOrValueExists = True
		Else
			blnKeyOrValueExists = False
		End If
		Call g_objFunctions.CreateDuplicateRecordset(m_rsEntries, rsEntries)

	End Function

	Public Function EnumRegistrySubkeysByKey(ByVal objRemoteRegServer, ByVal strConnectedWithThis, ByVal blnWMIRegGoodToGo, _
												ByVal blnRemoteRegGoodToGo, ByVal strRegistryHive, ByVal strRegistryKey, _
												ByVal blnIs64BitMachine, ByVal blnIsWow6432Node, ByVal strOSVersion, _
												ByRef blnKeyOrValueExists, ByRef rsKeys, ByRef objLogAndTrace, ByVal objLogAndTraceErrors, _
												ByRef objLogAndTraceLoadRS)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get all registry Subkeys using WMI Registry or RemoteRegistry processing
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					LogThis, DeleteAllRecordsetRows, EnumWMIRegistryProcessing, EnumRemoteRegistryKeys
	'*							ParseRegistryEnumToKeys, CreateDuplicateRecordset
	'*	Requirements:			Registry Constants
	'*****************************************************************************************************************************************
		Dim strKeyPath, blnExactMatch, blnSpecificKey, strSearchKey

		If (IsObject(objLogAndTrace)) Then
			Set m_objLogAndTrace = objLogAndTrace
		End If
		If (IsObject(objLogAndTraceErrors)) Then
			Set m_objLogAndTraceErrors = objLogAndTraceErrors
		End If
		If (IsObject(objLogAndTraceLoadRS)) Then
			Set m_objLogAndTraceLoadRS = objLogAndTraceLoadRS
		End If

		Call LogThis(vbTab & "Processing EnumRegistrySubkeysByKey started", m_objLogAndTrace)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsKeys)
		Call g_objFunctions.DeleteAllRecordsetRows(m_rsRegistryEnum)

		If (strRegistryKey = "") Then
			strKeyPath = ""
		Else
			'
			' Strip off the '\' at the end of the string if present
			'
			If (Right(strRegistryKey, 1) = "\") Then
				strKeyPath = Left(strRegistryKey, Len(strRegistryKey) - 1)
			Else
				strKeyPath = strRegistryKey
			End If
		End If
		blnIsWow6432Node = False
		blnExactMatch = False
		blnSpecificKey = False
		strSearchKey = ""
		'
		' The following functions load the m_rsRegistryEnum recordset with registry Key information
		'
		If (blnWMIRegGoodToGo) Then
			Call EnumWMIRegistryProcessing(objRemoteRegServer, strRegistryHive, strKeyPath, "EnumKey", blnIs64BitMachine, blnIsWow6432Node)
		Else
			Call EnumRemoteRegistryKeys(strConnectedWithThis, strRegistryHive, strKeyPath)
		End If
		'
		' Parse the m_rsRegistryEnum recordset for Keys that match
		'
		Call ParseRegistryEnumToKeys(blnIsWow6432Node, blnExactMatch, blnSpecificKey, strSearchKey)
		If (m_rsKeys.RecordCount > 0) Then
			blnKeyOrValueExists = True
		Else
			blnKeyOrValueExists = False
		End If
		Call g_objFunctions.CreateDuplicateRecordset(m_rsKeys, rsKeys)

	End Function

	Public Function SetRegistryEntry(ByVal objRemoteRegServer, ByVal strConnectedWithThis, ByVal blnWMIRegGoodToGo, _
										ByVal blnRemoteRegGoodToGo, ByVal strRegistryHive, ByVal strRegistryKey, ByVal strEntryName, _
										ByVal valToSet, ByVal strOperation, ByVal blnIs64BitMachine, ByVal blnIsWow6432Node, _
										ByRef objLogAndTrace, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Set the registry value using WMI Registry or RemoteRegistry processing
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				None
	'*  Calls:					LogThis, GetMethodToCall, SetWMIRegistryEntry, SetRemoteRegistryEntry
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strKeyPath, strMethodToCall, intRetVal, strDataType

		If (IsObject(objLogAndTrace)) Then
			Set m_objLogAndTrace = objLogAndTrace
		End If
		If (IsObject(objLogAndTraceErrors)) Then
			Set m_objLogAndTraceErrors = objLogAndTraceErrors
		End If

		Call LogThis(vbTab & "Processing SetRegistryKey started", objLogAndTrace)
		'
		' Strip off the '\' at the end of the string if present
		'
		If (Right(strRegistryKey, 1) = "\") Then
			strKeyPath = Left(strRegistryKey, Len(strRegistryKey) - 1)
		Else
			strKeyPath = strRegistryKey
		End If
		If (blnWMIRegGoodToGo) Then
			Call GetMethodToCall(strOperation, strMethodToCall)
			intRetVal = SetWMIRegistryEntry(objRemoteRegServer, strRegistryHive, strKeyPath, strEntryName, valToSet, strMethodToCall, _
												blnIs64BitMachine, blnIsWow6432Node)
		Else
			Call GetRegistryDataType(strOperation, strDataType)
			intRetVal = SetRemoteRegistryEntry(strConnectedWithThis, strRegistryHive, strRegistryKey, strEntryName, valToSet, _
													strDataType)
		End If
		SetRegistryEntry = intRetVal

	End Function

	Public Function DeleteRegistryEntry(ByVal objRemoteRegServer, ByVal strConnectedWithThis, ByVal blnWMIRegGoodToGo, _
										ByVal blnRemoteRegGoodToGo, ByVal strRegistryHive, ByVal strRegistryKey, ByVal strEntryName, _
										ByVal blnIs64BitMachine, ByVal blnIsWow6432Node, ByRef objLogAndTrace, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Delete the specified registry value using WMI Registry or RemoteRegistry processing
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				None
	'*  Calls:					LogThis, DeleteWMIRegistryEntry, DeleteRemoteRegistryEntry
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strKeyPath, strMethodToCall, intRetVal, strDataType

		If (IsObject(objLogAndTrace)) Then
			Set m_objLogAndTrace = objLogAndTrace
		End If
		If (IsObject(objLogAndTraceErrors)) Then
			Set m_objLogAndTraceErrors = objLogAndTraceErrors
		End If

		Call LogThis(vbTab & "Processing DeleteRegistryEntry started", objLogAndTrace)
		'
		' Strip off the '\' at the end of the string if present
		'
		If (Right(strRegistryKey, 1) = "\") Then
			strKeyPath = Left(strRegistryKey, Len(strRegistryKey) - 1)
		Else
			strKeyPath = strRegistryKey
		End If
		If (blnWMIRegGoodToGo) Then
			intRetVal = DeleteWMIRegistryEntry(objRemoteRegServer, strRegistryHive, strKeyPath, strEntryName, blnIs64BitMachine, blnIsWow6432Node)
		Else
			intRetVal = DeleteRemoteRegistryEntry(strConnectedWithThis, strRegistryHive, strRegistryKey, strEntryName)
		End If
		DeleteRegistryEntry = intRetVal

	End Function

	Public Function DeleteRegistryKey(ByVal objRemoteRegServer, ByVal strConnectedWithThis, ByVal blnWMIRegGoodToGo, _
										ByVal blnRemoteRegGoodToGo, ByVal strRegistryHive, ByVal strRegistryKey, ByVal blnIs64BitMachine, _
										ByVal blnIsWow6432Node, ByRef objLogAndTrace, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Delete the specified registry Key using WMI Registry or RemoteRegistry processing
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				None
	'*  Calls:					LogThis, DeleteWMIRegistryKey, DeleteRemoteRegistry
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strKeyPath, strMethodToCall, intRetVal, strDataType

		If (IsObject(objLogAndTrace)) Then
			Set m_objLogAndTrace = objLogAndTrace
		End If
		If (IsObject(objLogAndTraceErrors)) Then
			Set m_objLogAndTraceErrors = objLogAndTraceErrors
		End If

		Call LogThis(vbTab & "Processing DeleteRegistryKey started", objLogAndTrace)
		'
		' Strip off the '\' at the end of the string if present
		'
		If (Right(strRegistryKey, 1) = "\") Then
			strKeyPath = Left(strRegistryKey, Len(strRegistryKey) - 1)
		Else
			strKeyPath = strRegistryKey
		End If
		If (blnWMIRegGoodToGo) Then
			intRetVal = DeleteWMIRegistryKey(objRemoteRegServer, strRegistryHive, strKeyPath, blnIs64BitMachine, blnIsWow6432Node)
		Else
			intRetVal = DeleteRemoteRegistryKey(strConnectedWithThis, strRegistryHive, strRegistryKey)
		End If
		DeleteRegistryKey = intRetVal

	End Function

'#endregion

End Class

Class PossibleProcessing
	'
	' Requirements:		None
	'
	Private m_HKEY_LOCAL_MACHINE, m_REG_KEY_QUERY

	Private Sub Class_Initialize() 'Constructor
		If (IsObject(g_objFunctions) = False) Then
			WScript.Echo "Object g_objFunctions required for Class PossibleProcessing.  Abending..."
			WScript.Quit
		End If
		'
		' Registry Constants
		'
		m_HKEY_LOCAL_MACHINE = &H80000002
		m_REG_KEY_QUERY = &H0001
 	End Sub

	Private Sub Class_Terminate 'Destructor
    End Sub

	Private Sub LogThis(ByVal strText, ByRef objLogAndTrace)
		Dim strTextLocal
		If (IsObject(objLogAndTrace)) Then
			strTextLocal = strText
			Call g_objFunctions.CreatePrintableString(strTextLocal)
			objLogAndTrace.LogThis(strTextLocal)
		End If
	End Sub

	Private Function IsLocalAccountProcessingPossible(ByVal strConnectedWithThis, ByVal strLocalAdministratorGroupName, ByRef objLogAndTrace)
	'*****************************************************************************************************************************************
	'*  Purpose:				Check to see if remote registry calls are successful.
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				CanRegistryProcessingBeDone
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim objGroup, intErrNumber, strErrDescription
	
		Call LogThis("Checking access to Local Accounts on remote computer", objLogAndTrace)
		On Error Resume Next
		Set objGroup = GetObject("WinNT://" & strConnectedWithThis & "/" & strLocalAdministratorGroupName & ",group")
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
		
		If (intErrNumber = 0) Then
			Call LogThis("Access to Local Accounts on remote computer successful", objLogAndTrace)
			IsLocalAccountProcessingPossible = True
		Else
			Call LogThis("Access to Local Accounts on remote computer failed.  " & _
							"Error: " & intErrNumber & " (" & Hex(intErrNumber) & ")  Description: " & strErrDescription, objLogAndTrace)
			IsLocalAccountProcessingPossible = False
		End If
	
	End Function

	Private Function GetAdministratorsGroupName(ByRef objWMIServer, ByRef strLocalAdministratorGroupName)
	'*****************************************************************************************************************************************
	'*  Purpose:				Get the  to see if remote registry calls are successful.
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				CanRegistryProcessingBeDone
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strSQLQuery, intErrNumber, strErrDescription, colWMI, objWMI
		Const wbemFlagReturnWhenComplete = 0

		strSQLQuery = "SELECT * FROM Win32_Group WHERE LocalAccount = TRUE And SID = 'S-1-5-32-544'"
		Call g_objFunctions.ExecWMI(objWMIServer, intErrNumber, strErrDescription, colWMI, strSQLQuery, wbemFlagReturnWhenComplete, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
			For Each objWMI In colWMI
				strLocalAdministratorGroupName = objWMI.Name
			Next
		End If

	End Function

	Private Function GetLocalAccountProcessingSettings(ByVal strResolvedDNSHostName, ByVal strResolvedHostName, _
															ByVal strResolvedIPAddress, ByVal strResolvedNetBIOSName, _
															ByVal blnDNSHostNameResolved, ByVal blnHostNameResolved, _
															ByVal blnNetBIOSNameResolved, ByVal blnIPAddressResolved, _
															ByVal strLocalAdministratorGroupName,  ByRef strLocalAccountConnectedWithThis, _
															ByRef blnLocalAccountProcessingGoodToGo, ByRef objLogAndTrace)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determine what processing we can do
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				GatherMachineData
	'*  Calls:					IsLocalAccountProcessingPossible
	'*  Requirements:			None
	'*****************************************************************************************************************************************
	
		'
		' Check to see if Local Account processing is possible
		'
		If (blnIPAddressResolved) Then
			blnLocalAccountProcessingGoodToGo = IsLocalAccountProcessingPossible(strResolvedIPAddress, strLocalAdministratorGroupName, objLogAndTrace)
			If (blnLocalAccountProcessingGoodToGo) Then
				strLocalAccountConnectedWithThis = strResolvedIPAddress
				Exit Function
			End If
		End If
		If (blnDNSHostNameResolved) Then
			blnLocalAccountProcessingGoodToGo = IsLocalAccountProcessingPossible(strResolvedDNSHostName, strLocalAdministratorGroupName, objLogAndTrace)
			If (blnLocalAccountProcessingGoodToGo) Then
				strLocalAccountConnectedWithThis = strResolvedDNSHostName
				Exit Function
			End If
		End If
		If (blnHostNameResolved) Then
			blnLocalAccountProcessingGoodToGo = IsLocalAccountProcessingPossible(strResolvedHostName, strLocalAdministratorGroupName, objLogAndTrace)
			If (blnLocalAccountProcessingGoodToGo) Then
				strLocalAccountConnectedWithThis = strResolvedHostName
				Exit Function
			End If
		End If
		If (blnNetBIOSNameResolved) Then
			blnLocalAccountProcessingGoodToGo = IsLocalAccountProcessingPossible(blnNetBIOSNameResolved, strLocalAdministratorGroupName, objLogAndTrace)
			If (blnLocalAccountProcessingGoodToGo) Then
				strLocalAccountConnectedWithThis = blnNetBIOSNameResolved
				Exit Function
			End If
		End If
	
	End Function

	Private Function IsRemoteRegistryProcessingPossible(ByVal strConnectedWithThis, ByRef objLogAndTrace)
	'*****************************************************************************************************************************************
	'*  Purpose:				Check to see if remote registry calls are successful.
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				CanRegistryProcessingBeDone
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim objShell, strRegKey, objExec, strStandardError, blnStandardErrorSet, strStandardOut, blnStandardOutData
	
		Call LogThis("Checking remote registry access to HKLM\SYSTEM hive", objLogAndTrace)
		Set objShell = CreateObject("Wscript.Shell")
	
		strRegKey = "HKLM\SYSTEM"
		Set objExec = objShell.Exec("cmd /c REG QUERY \\" & strConnectedWithThis & "\" & strRegKey)
	
		strStandardError = objExec.StdErr.ReadAll
		blnStandardErrorSet = True
		If ((IsNull(strStandardError)) Or (IsEmpty(strStandardError)) Or (strStandardError = "") Or (strStandardError = " ")) Then
			blnStandardErrorSet = False
		End If
		
		strStandardOut = Trim(objExec.StdOut.ReadAll)
		If ((IsNull(strStandardOut)) Or (IsEmpty(strStandardOut)) Or (strStandardOut = "") Or (strStandardOut = " ")) Then
			blnStandardOutData = False
		Else
			blnStandardOutData = True
		End If
		'
		' Cleanup
		'
		Set objExec = Nothing
		Set objShell = Nothing
			
		If ((blnStandardErrorSet = False) And (blnStandardOutData)) Then
			Call LogThis("Remote registry processing IS possible", objLogAndTrace)
			IsRemoteRegistryProcessingPossible = True
		Else
			Call LogThis("Remote registry processing IS NOT possible", objLogAndTrace)
			IsRemoteRegistryProcessingPossible = False
		End If
	
	End Function

	Private Function GetRemoteRegistryProcessingSettings(ByVal strResolvedDNSHostName, ByVal strResolvedHostName, _
															ByVal strResolvedIPAddress, ByVal strResolvedNetBIOSName, _
															ByVal blnDNSHostNameResolved, ByVal blnHostNameResolved, _
															ByVal blnNetBIOSNameResolved, ByVal blnIPAddressResolved, _
															ByRef strConnectedWithThis, ByRef blnRemoteRegGoodToGo, _
															ByRef objLogAndTrace)
	'*****************************************************************************************************************************************
	'*  Purpose:				Determine what processing we can do
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				GatherMachineData
	'*  Calls:					IsRemoteRegistryProcessingPossible
	'*  Requirements:			None
	'*****************************************************************************************************************************************
	
		'
		' Check to see if RemoteRegistry processing is possible
		'
		If (blnIPAddressResolved) Then
			blnRemoteRegGoodToGo = IsRemoteRegistryProcessingPossible(strResolvedIPAddress, objLogAndTrace)
			If (blnRemoteRegGoodToGo) Then
				strConnectedWithThis = strResolvedIPAddress
				Exit Function
			End If
		End If
		If (blnDNSHostNameResolved) Then
			blnRemoteRegGoodToGo = IsRemoteRegistryProcessingPossible(strResolvedDNSHostName, objLogAndTrace)
			If (blnRemoteRegGoodToGo) Then
				strConnectedWithThis = strResolvedDNSHostName
				Exit Function
			End If
		End If
		If (blnHostNameResolved) Then
			blnRemoteRegGoodToGo = IsRemoteRegistryProcessingPossible(strResolvedHostName, objLogAndTrace)
			If (blnRemoteRegGoodToGo) Then
				strConnectedWithThis = strResolvedHostName
				Exit Function
			End If
		End If
		If (blnNetBIOSNameResolved) Then
			blnRemoteRegGoodToGo = IsRemoteRegistryProcessingPossible(blnNetBIOSNameResolved, objLogAndTrace)
			If (blnRemoteRegGoodToGo) Then
				strConnectedWithThis = blnNetBIOSNameResolved
				Exit Function
			End If
		End If
	
	End Function

	Private Function IsWMIRegistryProcessingPossible(ByVal objRemoteRegServer, ByRef objLogAndTrace)
	'*****************************************************************************************************************************************
	'*  Purpose:				Check to see if Win32 calls to the registry are successful with this connection.
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				GetWMIRegistryProcessingSettings
	'*  Calls:					CheckRegistryAccess
	'*	Requirements:			Registry Constants
	'*****************************************************************************************************************************************
		Dim blnRegistryAccessSystem, blnRegistryAccessSoftware, blnReturnValue
	
		Call LogThis("Checking registry access to SYSTEM and SOFTWARE hives", objLogAndTrace)
		blnRegistryAccessSystem = g_objFunctions.CheckRegistryAccess(objRemoteRegServer, m_HKEY_LOCAL_MACHINE, "SYSTEM", m_REG_KEY_QUERY)
		blnRegistryAccessSoftware = g_objFunctions.CheckRegistryAccess(objRemoteRegServer, m_HKEY_LOCAL_MACHINE, "SOFTWARE", m_REG_KEY_QUERY)
		If ((blnRegistryAccessSystem) And (blnRegistryAccessSoftware)) Then
			blnReturnValue = True
		Else
			blnReturnValue = False
		End If
		Call LogThis("Registry access to SYSTEM and SOFTWARE hives return value: " & blnReturnValue, objLogAndTrace)
		IsWMIRegistryProcessingPossible = blnReturnValue
	
	End Function

	Private Function IsWMIProcessingPossible(ByVal objWMIServer, ByRef objLogAndTrace)
	'*****************************************************************************************************************************************
	'*  Purpose:				Check to see if Win32 calls are successful with this connection.
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				CheckIfWMIProcessingBeDone
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strSQLQuery, colWMI, intErrNumber, strErrDescription, objWMI, blnGoodToProcessWMIComputerSystem, blnGoodToProcessWMIOperatingSystem
		Dim blnGoodToProcessWMI
		Const wbemFlagReturnWhenComplete = 0
		
		strSQLQuery = "SELECT Description,Manufacturer,BuildNumber,BuildType,Caption,Version FROM Win32_OperatingSystem"
		Call LogThis("Attempting WMI Query " & strSQLQuery, objLogAndTrace)
		Call g_objFunctions.ExecWMI(objWMIServer, intErrNumber, strErrDescription, colWMI, strSQLQuery, wbemFlagReturnWhenComplete, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
			Call LogThis("WMI Query to Win32_OperatingSystem Successful", objLogAndTrace)
			blnGoodToProcessWMIOperatingSystem = True
		Else
			Call LogThis("WMI Query to Win32_OperatingSystem Failed", objLogAndTrace)
			blnGoodToProcessWMIOperatingSystem = False
		End If

		strSQLQuery = "SELECT Name,DomainRole,Domain FROM Win32_ComputerSystem"
		Call LogThis("Attempting WMI Query " & strSQLQuery, objLogAndTrace)
		Call g_objFunctions.ExecWMI(objWMIServer, intErrNumber, strErrDescription, colWMI, strSQLQuery, wbemFlagReturnWhenComplete, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
			Call LogThis("WMI Query to Win32_ComputerSystem Successful", objLogAndTrace)
			blnGoodToProcessWMIComputerSystem = True
		Else
			Call LogThis("WMI Query to Win32_ComputerSystem Failed", objLogAndTrace)
			blnGoodToProcessWMIComputerSystem = False
		End If
		
		If ((blnGoodToProcessWMIComputerSystem) And (blnGoodToProcessWMIOperatingSystem)) Then
			blnGoodToProcessWMI = True
		Else
			blnGoodToProcessWMI = False
		End If
		Call LogThis("IsWMIProcessingPossible return value: " & blnGoodToProcessWMI, objLogAndTrace)
		IsWMIProcessingPossible = blnGoodToProcessWMI
	
	End Function

	Private Function MakeCOMRegistryConnection(ByVal strPassedParam, ByVal blnProcessingLocal, ByVal strUserID, ByVal strPassword, _
													ByVal strResolvedDNSHostName, ByVal strResolvedHostName, ByVal strResolvedIPAddress, _
													ByVal strResolvedNetBIOSName, ByVal blnDNSHostNameResolved, ByVal blnHostNameResolved, _
													ByVal blnNetBIOSNameResolved, ByVal blnIPAddressResolved, ByVal strMyOSVersion, _
													ByRef strWMIRegistryConnectedWithThis, ByRef objRemoteRegServer, _
													ByRef strWMIRegConnectError, ByRef objLogAndTrace, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Check to see if a COM connection can be made
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					CreateServerRegistryConnection
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim intRetVal, intErrNumber, strErrDescription, strError
	
		If (blnIPAddressResolved) Then
			Call LogThis("Attempting WMI Registry connection to IP Address " & strResolvedIPAddress, objLogAndTrace)
			intRetVal = g_objFunctions.CreateServerRegistryConnection(strResolvedIPAddress, objRemoteRegServer, intErrNumber, strErrDescription, _
																		strError, strUserID, strPassword, objLogAndTraceErrors)
			If (intRetVal = 0) Then
				Call LogThis("WMI Registry connection to IP Address " & strResolvedIPAddress & " successful", objLogAndTrace)
				strWMIRegistryConnectedWithThis = strResolvedIPAddress
				strWMIRegConnectError = ""
				MakeCOMRegistryConnection = True
				Exit Function
			End If
			If (strWMIRegConnectError = "") Then
				strWMIRegConnectError = strError
			Else
				strWMIRegConnectError = "    " & strError
			End If
		End If
		If (blnDNSHostNameResolved) Then
			Call LogThis("Attempting WMI Registry FQDN connection to " & strResolvedDNSHostName, objLogAndTrace)
			intRetVal = g_objFunctions.CreateServerRegistryConnection(strResolvedDNSHostName, objRemoteRegServer, intErrNumber, strErrDescription, _
																		strError, strUserID, strPassword, objLogAndTraceErrors)
			If (intRetVal = 0) Then
				Call LogThis("WMI Registry FQDN connection to " & strResolvedDNSHostName & " successful", objLogAndTrace)
				strWMIRegistryConnectedWithThis = strResolvedDNSHostName
				strWMIRegConnectError = ""
				MakeCOMRegistryConnection = True
				Exit Function
			End If
			If (strWMIRegConnectError = "") Then
				strWMIRegConnectError = strError
			Else
				strWMIRegConnectError = "    " & strError
			End If
		End If
		If (blnHostNameResolved) Then
			Call LogThis("Attempting WMI Registry HostName connection to " & strResolvedHostName, objLogAndTrace)
			intRetVal = g_objFunctions.CreateServerRegistryConnection(strResolvedHostName, objRemoteRegServer, intErrNumber, strErrDescription, _
																		strError, strUserID, strPassword, objLogAndTraceErrors)
			If (intRetVal = 0) Then
				Call LogThis("WMI Registry HostName connection to " & strResolvedHostName & " successful", objLogAndTrace)
				strWMIRegistryConnectedWithThis = strResolvedHostName
				strWMIRegConnectError = ""
				MakeCOMRegistryConnection = True
				Exit Function
			End If
			If (strWMIRegConnectError = "") Then
				strWMIRegConnectError = strError
			Else
				strWMIRegConnectError = "    " & strError
			End If
		End If
		If (blnNetBIOSNameResolved) Then
			Call LogThis("Attempting WMI Registry NetBIOS connection to " & strResolvedNetBIOSName, objLogAndTrace)
			intRetVal = g_objFunctions.CreateServerRegistryConnection(strResolvedNetBIOSName, objRemoteRegServer, intErrNumber, strErrDescription, _
																		strError, strUserID, strPassword, objLogAndTraceErrors)
			If (intRetVal = 0) Then
				Call LogThis("WMI Registry NetBIOS connection to " & strResolvedNetBIOSName & " successful", objLogAndTrace)
				strWMIRegistryConnectedWithThis = strResolvedNetBIOSName
				strWMIRegConnectError = ""
				MakeCOMRegistryConnection = True
				Exit Function
			End If
			If (strWMIRegConnectError = "") Then
				strWMIRegConnectError = strError
			Else
				strWMIRegConnectError = "    " & strError
			End If
		End If
		'
		' Try to connect with whatever was passed
		'
		Call LogThis("Attempting WMI connection (via passed parameter) to " & strPassedParam, objLogAndTrace)
		intRetVal = g_objFunctions.CreateServerRegistryConnection(strPassedParam, objRemoteRegServer, intErrNumber, strErrDescription, strError, _
																	strUserID, strPassword, objLogAndTraceErrors)
		If (intRetVal = 0) Then
			Call LogThis("WMI connection (via passed parameter) to " & strPassedParam & " successful", objLogAndTrace)
			strWMIRegistryConnectedWithThis = strPassedParam
			strWMIRegConnectError = ""
			MakeCOMRegistryConnection = True
			Exit Function
		End If
		If (strWMIRegConnectError = "") Then
			strWMIRegConnectError = strError
		Else
			strWMIRegConnectError = "    " & strError
		End If
		'
		' No good so WMI will not be available
		'
		Call LogThis("WMI registry connection to " & strPassedParam & " failed", objLogAndTrace)
		strWMIRegistryConnectedWithThis = ""
		MakeCOMRegistryConnection = False
	
	End Function

	Public Function ValidateWMIRegistryOpportunities(ByVal strPassedParam, ByVal blnProcessingLocal, ByVal strUserID, _
														ByVal strPassword, ByVal strResolvedDNSHostName, ByVal strResolvedHostName, _
														ByVal strResolvedIPAddress, ByVal strResolvedNetBIOSName, _
														ByVal blnDNSHostNameResolved, ByVal blnHostNameResolved, _
														ByVal blnNetBIOSNameResolved, ByVal blnIPAddressResolved, ByRef blnWMIRegGoodToGo, _
														ByRef strWMIRegistryConnectedWithThis, ByRef objRemoteRegServer, _
														ByRef strWMIRegConnectError, ByRef blnWMIRegConnectErrorOccurred, _
														ByRef objLogAndTrace, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Validate WMI Registry processing possibilities
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					MakeCOMRegistryConnection, IsWMIRegistryProcessingPossible
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strMyOSVersion, blnCOMRegistryConnectionSuccessful
	
		strMyOSVersion = g_objFunctions.GetOSVersion()
		blnWMIRegGoodToGo = False
		blnCOMRegistryConnectionSuccessful = MakeCOMRegistryConnection(strPassedParam, blnProcessingLocal, strUserID, strPassword, _
																			strResolvedDNSHostName, strResolvedHostName, _
																			strResolvedIPAddress, strResolvedNetBIOSName, _
																			blnDNSHostNameResolved, blnHostNameResolved, _
																			blnNetBIOSNameResolved, blnIPAddressResolved, strMyOSVersion, _
																			strWMIRegistryConnectedWithThis, objRemoteRegServer, _
																			strWMIRegConnectError, objLogAndTrace, objLogAndTraceErrors)
		If (blnCOMRegistryConnectionSuccessful) Then
			blnWMIRegGoodToGo = IsWMIRegistryProcessingPossible(objRemoteRegServer, objLogAndTrace)
		End If
		If (strWMIRegConnectError <> "") Then
			blnWMIRegConnectErrorOccurred = True
		Else
			blnWMIRegConnectErrorOccurred = False
		End If

	End Function

	Private Function MakeCOMConnection(ByVal strPassedParam, ByVal strUserID, ByVal strPassword, ByVal strResolvedDNSHostName, _
											ByVal strResolvedHostName, ByVal strResolvedIPAddress, ByVal strResolvedNetBIOSName, _
											ByVal blnDNSHostNameResolved, ByVal blnHostNameResolved, ByVal blnNetBIOSNameResolved, _
											ByVal blnIPAddressResolved, ByRef strConnectedWithThis, ByRef objWMIServer, _
											ByRef strWMIConnectError, ByRef objLogAndTrace, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Check to see if a COM connection can be made
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					CreateServerConnection
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strNamespace, intRetVal, intErrNumber, strErrDescription, strError
	
		strNamespace = "root\cimv2"
		'
		' Try to connect with anything returned from GetClientNetworkInfo
		'
		If (blnIPAddressResolved) Then
			Call LogThis("Attempting WMI connection to IP Address " & strResolvedIPAddress, objLogAndTrace)
			intRetVal = g_objFunctions.CreateServerConnection(strResolvedIPAddress, objWMIServer, intErrNumber, strErrDescription, _
																strError, strNamespace, strUserID, strPassword, objLogAndTraceErrors)
			If (intRetVal = 0) Then
				Call LogThis("WMI connection to IP Address " & strResolvedIPAddress & " successful", objLogAndTrace)
				strConnectedWithThis = strResolvedIPAddress
				strWMIConnectError = ""
				MakeCOMConnection = True
				Exit Function
			End If
			If (strWMIConnectError = "") Then
				strWMIConnectError = strError
			Else
				strWMIConnectError = "    " & strError
			End If
		End If
		If (blnDNSHostNameResolved) Then
			Call LogThis("Attempting WMI FQDN connection to " & strResolvedDNSHostName, objLogAndTrace)
			intRetVal = g_objFunctions.CreateServerConnection(strResolvedDNSHostName, objWMIServer, intErrNumber, strErrDescription, _
																strError, strNamespace, strUserID, strPassword, objLogAndTraceErrors)
			If (intRetVal = 0) Then
				Call LogThis("WMI FQDN connection to " & strResolvedDNSHostName & " successful", objLogAndTrace)
				strConnectedWithThis = strResolvedDNSHostName
				strWMIConnectError = ""
				MakeCOMConnection = True
				Exit Function
			End If
			If (strWMIConnectError = "") Then
				strWMIConnectError = strError
			Else
				strWMIConnectError = "    " & strError
			End If
		End If
		If (blnHostNameResolved) Then
			Call LogThis("Attempting WMI HostName connection to " & strResolvedHostName, objLogAndTrace)
			intRetVal = g_objFunctions.CreateServerConnection(strResolvedHostName, objWMIServer, intErrNumber, strErrDescription, _
																strError, strNamespace, strUserID, strPassword, objLogAndTraceErrors)
			If (intRetVal = 0) Then
				Call LogThis("WMI HostName connection to " & strResolvedHostName & " successful", objLogAndTrace)
				strConnectedWithThis = strResolvedHostName
				strWMIConnectError = ""
				MakeCOMConnection = True
				Exit Function
			End If
			If (strWMIConnectError = "") Then
				strWMIConnectError = strError
			Else
				strWMIConnectError = "    " & strError
			End If
		End If
		If (blnNetBIOSNameResolved) Then
			Call LogThis("Attempting WMI NetBIOS connection to " & strResolvedNetBIOSName, objLogAndTrace)
			intRetVal = g_objFunctions.CreateServerConnection(strResolvedNetBIOSName, objWMIServer, intErrNumber, strErrDescription, _
																strError, strNamespace, strUserID, strPassword, objLogAndTraceErrors)
			If (intRetVal = 0) Then
				Call LogThis("WMI NetBIOS connection to " & strResolvedNetBIOSName & " successful", objLogAndTrace)
				strConnectedWithThis = strResolvedNetBIOSName
				strWMIConnectError = ""
				MakeCOMConnection = True
				Exit Function
			End If
			If (strWMIConnectError = "") Then
				strWMIConnectError = strError
			Else
				strWMIConnectError = "    " & strError
			End If
		End If
		'
		' Try to connect with whatever was passed
		'
		Call LogThis("Attempting WMI connection (via passed parameter) to " & strPassedParam, objLogAndTrace)
		intRetVal = g_objFunctions.CreateServerConnection(strPassedParam, objWMIServer, intErrNumber, strErrDescription, strError, _
															strNamespace, strUserID, strPassword, objLogAndTraceErrors)
		If (intRetVal = 0) Then
			Call LogThis("WMI connection (via passed parameter) to " & strPassedParam & " successful", objLogAndTrace)
			strConnectedWithThis = strPassedParam
			MakeCOMConnection = True
			Exit Function
		End If
		If (strWMIConnectError = "") Then
			strWMIConnectError = strError
		Else
			strWMIConnectError = "    " & strError
		End If
		'
		' No good so WMI will not be available
		'
		Call LogThis("All WMI connection attempts failed", objLogAndTrace)
		strConnectedWithThis = ""
		MakeCOMConnection = False
	
	End Function

	Public Function ValidateWMIOpportunities(ByVal strUserID, ByVal strPassword, ByVal strPassedParam, ByVal strResolvedDNSHostName, _
												ByVal strResolvedHostName, ByVal strResolvedIPAddress, ByVal strResolvedNetBIOSName, _
												ByVal blnDNSHostNameResolved, ByVal blnHostNameResolved, ByVal blnNetBIOSNameResolved, _
												ByVal blnIPAddressResolved, ByRef strWMIConnectedWithThis, ByRef objRemoteWMIServer, _
												ByRef blnWMIGoodToGo, ByRef blnCOMConnectionSuccessful, ByRef strWMIConnectError, _
												ByRef blnWMIConnectErrorOccurred, ByRef objLogAndTrace, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Validate WMI processing possibilities
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					MakeCOMConnection, IsWMIProcessingPossible
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		strWMIConnectError = ""
		'
		' Determine what processing is possible - connect via WMI (DCOM) to machine
		'
		blnCOMConnectionSuccessful = MakeCOMConnection(strPassedParam, strUserID, strPassword, strResolvedDNSHostName, _
															strResolvedHostName, strResolvedIPAddress, strResolvedNetBIOSName, _
															blnDNSHostNameResolved, blnHostNameResolved, blnNetBIOSNameResolved, _
															blnIPAddressResolved, strWMIConnectedWithThis, objRemoteWMIServer, _
															strWMIConnectError, objLogAndTrace, objLogAndTraceErrors)
		If (blnCOMConnectionSuccessful) Then
			'
			' COM connection was successful - see if we can access WMI
			'
			blnWMIGoodToGo = IsWMIProcessingPossible(objRemoteWMIServer, objLogAndTrace)
		End If
		If (strWMIConnectError <> "") Then
			blnWMIConnectErrorOccurred = True
		Else
			blnWMIConnectErrorOccurred = False
		End If

	End Function

	Public Function ValidateProcessingOpportunities(ByVal strPassedParam, ByVal strUserID, ByVal strPassword, _
														ByVal strResolvedDNSHostName, ByVal strResolvedHostName, _
														ByVal strResolvedIPAddress, ByVal strResolvedNetBIOSName, _
														ByVal blnDNSHostNameResolved, ByVal blnHostNameResolved, _
														ByVal blnNetBIOSNameResolved, ByVal blnIPAddressResolved, ByVal blnProcessingLocal, _
														ByRef strWMIConnectedWithThis, ByRef objRemoteWMIServer, _
														ByRef blnWMIGoodToGo, ByRef strWMIRegistryConnectedWithThis, _
														ByRef objRemoteRegServer, ByRef blnWMIRegGoodToGo, _
														ByRef strRemoteRegConnectedWithThis, ByRef blnRemoteRegGoodToGo, _
														ByRef strLocalAccountsConnectedWithThis, ByRef blnLocalAccountProcessingGoodToGo, _
														ByRef blnWMIConnectErrorOccurred, ByRef strWMIConnectError, _
														ByRef blnWMIRegConnectErrorOccurred, ByRef strWMIRegConnectError, _
														ByRef blnCOMConnectionSuccessful, ByRef objLogAndTrace, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Validate all processing possibilities
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					MakeCOMConnection, IsWMIProcessingPossible, MakeCOMRegistryConnection, IsWMIRegistryProcessingPossible
	'*							GetRemoteRegistryProcessingSettings, GetLocalAccountProcessingSettings
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim strMyOSVersion, blnCOMRegistryConnectionSuccessful, strLocalAdministratorGroupName
	
		'
		' WMI
		'
		Call ValidateWMIOpportunities(strUserID, strPassword, strPassedParam, strResolvedDNSHostName, strResolvedHostName, _
										strResolvedIPAddress, strResolvedNetBIOSName, blnDNSHostNameResolved, blnHostNameResolved, _
										blnNetBIOSNameResolved, blnIPAddressResolved, strWMIConnectedWithThis, objRemoteWMIServer, _
										blnWMIGoodToGo, blnCOMConnectionSuccessful, strWMIConnectError, blnWMIConnectErrorOccurred, _
										objLogAndTrace, objLogAndTraceErrors)
		'
		' Initialize starting point for connection attempts
		'
		strWMIRegistryConnectedWithThis = strWMIConnectedWithThis
		strRemoteRegConnectedWithThis = strWMIConnectedWithThis
		strLocalAccountsConnectedWithThis = strWMIConnectedWithThis
		'
		' WMI Registry
		'
		Call ValidateWMIRegistryOpportunities(strPassedParam, blnProcessingLocal, strUserID, strPassword, strResolvedDNSHostName, _
												strResolvedHostName, strResolvedIPAddress, strResolvedNetBIOSName, blnDNSHostNameResolved, _
												blnHostNameResolved, blnNetBIOSNameResolved, blnIPAddressResolved, blnWMIRegGoodToGo, _
												strWMIRegistryConnectedWithThis, objRemoteRegServer, strWMIRegConnectError, _
												blnWMIRegConnectErrorOccurred, objLogAndTrace, objLogAndTraceErrors)
		'
		' Get Remote Registry Processing Settings
		'
		Call GetRemoteRegistryProcessingSettings(strResolvedDNSHostName, strResolvedHostName, strResolvedIPAddress, strResolvedNetBIOSName, _
													blnDNSHostNameResolved, blnHostNameResolved, blnNetBIOSNameResolved, blnIPAddressResolved, _
													strRemoteRegConnectedWithThis, blnRemoteRegGoodToGo, objLogAndTrace)
		'
		' Get Local Administator Group name on Remote machine (it's not always Administrators)
		'
		strLocalAdministratorGroupName = "Administrators"
		If (blnWMIGoodToGo) Then
			Call GetAdministratorsGroupName(objRemoteWMIServer, strLocalAdministratorGroupName)
		End If
		'
		' Get Local Accounts Processing Settings
		'
		Call GetLocalAccountProcessingSettings(strResolvedDNSHostName, strResolvedHostName, strResolvedIPAddress, strResolvedNetBIOSName, _
													blnDNSHostNameResolved, blnHostNameResolved, blnNetBIOSNameResolved, blnIPAddressResolved, _
													strLocalAdministratorGroupName, strLocalAccountsConnectedWithThis, _
													blnLocalAccountProcessingGoodToGo, objLogAndTrace)
	End Function
	
End Class

Class XMLProcessing
	'
	' Requirements:		None
	'
	
	Private Sub Class_Initialize() 'Constructor
		If (IsObject(g_objFunctions) = False) Then
			WScript.Echo "Object g_objFunctions required for Class PossibleProcessing.  Abending..."
			WScript.Quit
		End If
 	End Sub

	Private Sub Class_Terminate 'Destructor
    End Sub

	Public Function CreateSetAndLinkAttribute(ByRef xmlDoc, ByRef xmlAttributeNode, ByVal strAttribute, ByVal strAttributeValue)
	'*****************************************************************************************************************************************
	'*  Purpose:				Creates the Database and associated tables (if necessary).
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				CreateFileChecksEntity, CreateRegistryChecksEntity, BuildXMLOutputForGMDProcessing
	'*  Calls:					StripNull
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim xmlAttribute
		
		Set xmlAttribute = xmlDoc.createAttribute(strAttribute)
		'
		' Set the value of the attribute
		'
		xmlAttribute.text = g_objFunctions.StripNull(strAttributeValue)
		'
		' Append the id attribute to the field element
		'
		xmlAttributeNode.setAttributeNode xmlAttribute
	
	End Function

	Public Function AppendChild(ByRef xmlElementChild, ByRef xmlElementAppendChild)
	'*****************************************************************************************************************************************
	'*  Purpose:				Creates the Database and associated tables (if necessary).
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				CreateFileChecksEntity, CreateRegistryChecksEntity, BuildXMLOutputForGMDProcessing
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
	
		xmlElementAppendChild.appendChild xmlElementChild
	
	End Function

	Public Function AppendChildAndSetText(ByRef xmlDoc, ByVal strElement, ByRef xmlElementAppendChild, ByVal strText)
	'*****************************************************************************************************************************************
	'*  Purpose:				Creates the Database and associated tables (if necessary).
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				CreateFileChecksEntity, CreateRegistryChecksEntity, BuildXMLOutputForGMDProcessing
	'*  Calls:					StripNull
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim xmlElementChild
	
		Set xmlElementChild = xmlDoc.createElement(strElement)
		xmlElementChild.Text = g_objFunctions.StripNull(strText)
		xmlElementAppendChild.appendChild xmlElementChild
	
	End Function

	Public Function CreateNewElement(ByRef xmlDoc, ByRef xmlElementToCreate, ByVal strElementName)
	'*****************************************************************************************************************************************
	'*  Purpose:				Creates a new field element under an existing Recordset element
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					AppendChild, CreateSetAndLinkAttribute
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		
		Set xmlElementToCreate = xmlDoc.createElement(strElementName)
	
	End Function

	Public Function CreateAppendNewElement(ByRef xmlDoc, ByVal xmlElement, ByRef xmlElementToCreate, ByVal strElementName)
	'*****************************************************************************************************************************************
	'*  Purpose:				Creates a new field element under an existing Recordset element
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					AppendChild, CreateSetAndLinkAttribute
	'*	Requirements:			None
	'*****************************************************************************************************************************************

		Set xmlElementToCreate = xmlDoc.createElement(strElementName)
		Call AppendChild(xmlElementToCreate, xmlElement)
	
	End Function

	Public Function CreateNewTableElement(ByRef xmlDoc, ByRef xmlElement, ByVal strTableName, ByVal strDatabaseName, ByVal strSchemaName)
	'*****************************************************************************************************************************************
	'*  Purpose:				Creates a new table element
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				All
	'*  Calls:					CreateSetAndLinkAttribute
	'*	Requirements:			None
	'*****************************************************************************************************************************************

		Set xmlElement = xmlDoc.createElement("Table")
		Call CreateSetAndLinkAttribute(xmlDoc, xmlElement, "TableName", strTableName)
		Call CreateSetAndLinkAttribute(xmlDoc, xmlElement, "SchemaName", strSchemaName)
		Call CreateSetAndLinkAttribute(xmlDoc, xmlElement, "DatabaseName", strDatabaseName)
	
	End Function

	Public Function CreateAndOpenTableRecordset(ByRef rsToCreate)
	'*****************************************************************************************************************************************
	'*  Purpose:				Create the Table recordset
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Any
	'*  Calls:					DeleteAllRecordsetRows (library)
	'*  Requirements:			Generic Constants (library)
	'*****************************************************************************************************************************************

		If (rsToCreate.Fields.Count = 0) Then
			rsToCreate.Fields.Append "TableName", m_adVarChar, 50
			rsToCreate.Fields.Append "ColumnName", m_adVarChar, 64
			rsToCreate.Fields.Append "ColumnType", m_adVarChar, 10
			rsToCreate.Fields.Append "MaxSize", m_adInteger
			rsToCreate.Open
		Else
			'
			' Delete any existing records
			'
			Call g_objFunctions.DeleteAllRecordsetRows(rsToCreate)
			If (rsToCreate.State <> m_adStateOpen) Then
				rsToCreate.Open
			End If
		End If

	End Function

	Public Function ProcessTableXML(ByVal strXMLFile, ByRef rsTable, ByVal strRequestCriteria, ByRef objLogAndTraceLoadRS, _
										ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Open and process the Table entries from the XML file
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					CreateAndOpenTableRecordset, LoadRS
	'*  Requirements:			Global Constants
	'*****************************************************************************************************************************************
		Dim xmlDoc, strCriteria, objNodeList, objRowNodes, objNode, objRowNode
		Dim strColumnName, strColumnType, intMaxSize

		Call CreateAndOpenTableRecordset(rsTable)
		'
		' Create a new XML Document object and load the XML code into it.
		'
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
		xmlDoc.async = False					' Load the entire document before continuing processing
		xmlDoc.preserveWhiteSpace = True
		xmlDoc.resolveExternals = False
		'
		' Use the next 2 commands to perform a case-insensitive search
		'
		xmlDoc.setProperty "SelectionLanguage", "XPath"
		xmlDoc.setProperty "SelectionNamespaces", "xmlns:ms='urn:schemas-microsoft-com:xslt'"
		xmlDoc.validateOnParse = False
		xmlDoc.load strXMLFile

		strCriteria = "@tablename='" & strRequestCriteria & "'"
		Set objNodeList = xmlDoc.selectNodes("//table[" & strCriteria & "]")
		Set objRowNodes = xmlDoc.selectNodes("//table[" & strCriteria & "]/row")

		For Each objNode In objNodeList
			For Each objRowNode In objRowNodes
				strColumnName = objRowNode.SelectSingleNode("columnname").Text
				strColumnType = objRowNode.SelectSingleNode("columntype").Text
				intMaxSize = objRowNode.SelectSingleNode("maxsize").Text
				'
				' The variables are loaded - load the recordset
				'
				rsTable.AddNew
				Call g_objFunctions.LoadRS(rsTable, "TableName", strRequestCriteria, objLogAndTraceLoadRS, objLogAndTraceErrors)
				Call g_objFunctions.LoadRS(rsTable, "ColumnName", strColumnName, objLogAndTraceLoadRS, objLogAndTraceErrors)
				Call g_objFunctions.LoadRS(rsTable, "ColumnType", strColumnType, objLogAndTraceLoadRS, objLogAndTraceErrors)
				Call g_objFunctions.LoadRS(rsTable, "MaxSize", intMaxSize, objLogAndTraceLoadRS, objLogAndTraceErrors)
				rsTable.Update
			Next
		Next

	End Function

	Public Function ConvertSQLColumnTypeAndSizeToRecordset(ByVal strColumnType, ByRef strRecordsetColumnType, ByVal strMaxSize, _
																ByRef intMaxSize)
	'*****************************************************************************************************************************************
	'*  Purpose:				Converts the information passed into datatypes that are compatible with recordset creation
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				CreateDBObjectTableEntity
	'*  Calls:					None
	'*  Requirements:			None
	'*****************************************************************************************************************************************

		Select Case UCase(strColumnType)
			Case "VARCHAR"
				strRecordsetColumnType = "TEXT"
				If (IsNumeric(strMaxSize)) Then
					intMaxSize = CInt(strMaxSize)
				Else
					intMaxSize = -1
				End If
			Case "TINYINT"
				strRecordsetColumnType = "TINYINT"
			Case "SMALLINT"
				strRecordsetColumnType = "SMALLINT"
			Case "INTEGER"
				strRecordsetColumnType = "INT"
			Case "BIT"
				strRecordsetColumnType = "BOOLEAN"
			Case "FLOAT"
				strRecordsetColumnType = "FLOAT"
			Case "SMALLDATETIME"
				strRecordsetColumnType = "DATETIME"
		End Select

	End Function

	Public Function ConvertTableXMLToRecordset(ByVal strXMLFile, ByRef rsTableInfo, ByRef objLogAndTraceLoadRS, ByRef objLogAndTraceErrors)
	'*****************************************************************************************************************************************
	'*  Purpose:				Process existing XML file and extract the "Create Table" section into a recordset
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					LoadRS
	'*  Requirements:			Global Constants
	'*****************************************************************************************************************************************
		Dim xmlDoc, strCriteria, objNodeList, objNode, strCDATA, strCreateTable
		Dim strTableName, strWork, arrRows, intCount, strColumnName, strColumnType
		Dim strColumnLength, intColumnLength, strConvertedColumnType

		Call CreateAndOpenTableRecordset(rsTableInfo)
		'
		' Create a new XML Document object and load the XML code into it.
		'
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
		xmlDoc.async = False					' Load the entire document before continuing processing
		xmlDoc.preserveWhiteSpace = True
		xmlDoc.resolveExternals = False
		xmlDoc.setProperty "SelectionLanguage", "XPath"
		xmlDoc.setProperty "SelectionNamespaces", "xmlns:ms='urn:schemas-microsoft-com:xslt'"
		xmlDoc.validateOnParse = False
		xmlDoc.load strXMLFile
		'
		' The following code pulls the CDATA text from all of the Create table Elements
		'
		strCriteria = "objecttype='Table'"
		Set objNodeList = xmlDoc.SelectNodes("//dbobject[" & strCriteria & "]")
		For Each objNode In objNodeList
			strCDATA = objNode.SelectSingleNode("sql").Text
			'
			' Replace Tabs and Line Feeds (the vbCrLf was the only one I found)
			'
			strCDATA = Replace(strCDATA, vbTab, "")
			strCDATA = Replace(strCDATA, vbCrLf, "")
			strCDATA = Replace(strCDATA, vbCr, "")
			strCDATA = Replace(strCDATA, vbLf, "")
			'
			' Parse the "Create Table" command
			'
			strCreateTable = Split(strCDATA, "CREATE TABLE", 2, vbTextCompare)(1)
			strCreateTable = Split(strCreateTable, ";", 2)(0)
			strTableName = Trim(Split(strCreateTable, "(", 2)(0))
			strWork = Trim(Split(strCreateTable, "(", 2)(1))
			arrRows = Split(strWork, ",")
			For intCount = 0 To UBound(arrRows)
				strColumnName = Trim(Split(arrRows(intCount), " ", 3)(0))
				strColumnType = Trim(Split(arrRows(intCount), " ", 3)(1))
				strColumnLength = ""
				If (InStr(strColumnType, "(") And InStr(strColumnType, ")")) Then
					strColumnLength = Replace(Split(strColumnType, "(", 2)(1), ")", "")
					strColumnType = Split(strColumnType, "(", 2)(0)
				End If
				If (InStr(strColumnType, ")")) Then
					strColumnType = Replace(strColumnType, ")", "")
				End If
				Call ConvertSQLColumnTypeAndSizeToRecordset(strColumnType, strConvertedColumnType, strColumnLength, intColumnLength)			
				'
				' Load the recordset
				'
				rsTableInfo.AddNew
				Call g_objFunctions.LoadRS(rsTableInfo, "TableName", strTableName, objLogAndTraceLoadRS, objLogAndTraceErrors)
				Call g_objFunctions.LoadRS(rsTableInfo, "ColumnName", strColumnName, objLogAndTraceLoadRS, objLogAndTraceErrors)
				Call g_objFunctions.LoadRS(rsTableInfo, "ColumnType", strColumnType, objLogAndTraceLoadRS, objLogAndTraceErrors)
				Call g_objFunctions.LoadRS(rsTableInfo, "MaxSize", intColumnLength, objLogAndTraceLoadRS, objLogAndTraceErrors)
				rsTableInfo.Update
			Next
		Next

	End Function

	Public Function FormatDateForXMLOutput()
	'*****************************************************************************************************************************************
	'*  Purpose:				Creates an output date formatted like yyyy-mm-dd
	'*  Arguments supplied:		Look up
	'*  Return Value:			0 to indicate success
	'*  Called by:				BuildXMLOutputForGMDProcessing
	'*  Calls:					None
	'*	Requirements:			None
	'*****************************************************************************************************************************************
		Dim dtCurrent, strYear, strMonth, strDay

		dtCurrent = Now()
		strYear = DatePart("YYYY", dtCurrent)
		strMonth = DatePart("M", dtCurrent)
		If (Len(strMonth) < 2) Then
			strMonth = "0" & strMonth
		End If
		strDay = DatePart("D", dtCurrent)
		If (Len(strDay) < 2) Then
			strDay = "0" & strDay
		End If
		FormatDateForXMLOutput = strYear & "-" & strMonth & "-" & strDay

	End Function

	Public Function FormatAndSaveXMLFile(ByVal strOutputFile, ByRef xmlElementTables, ByRef xmlElementPassedParams, _
											ByRef xmlElementProcessingInfo)
	'*****************************************************************************************************************************************
	'*  Purpose:				Setup all database recordsets for processing
	'*  Arguments supplied:		None
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					CreateAppendNewElement, CreateNewElement, CreateSetAndLinkAttribute, AppendChild
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim xmlDocLocal, xmlElementRemoteData, xmlProcessingInstruction

		Set xmlDocLocal = CreateObject("Microsoft.XMLDOM")
		Call CreateAppendNewElement(xmlDocLocal, xmlDocLocal, xmlElementRemoteData, "RemoteData")
		'
		' Append PassedParams, ProcessingInfo, and Tables to the RemoteData Element
		'
		Call g_objXMLProcessing.AppendChild(xmlElementPassedParams, xmlElementRemoteData)
		Call g_objXMLProcessing.AppendChild(xmlElementProcessingInfo, xmlElementRemoteData)
		Call g_objXMLProcessing.AppendChild(xmlElementTables, xmlElementRemoteData)
		'
		' Format the XML file
		'
		Set xmlProcessingInstruction = xmlDocLocal.createProcessingInstruction("xml","version='1.0' encoding='utf-8'")
		xmlDocLocal.insertBefore xmlProcessingInstruction, xmlDocLocal.childNodes(0)
		'
		' Save the XML file
		'
		xmlDocLocal.save strOutputFile
		'
		' Cleanup
		'
		Set xmlDocLocal = Nothing
		Set xmlElementRemoteData = Nothing
		Set xmlProcessingInstruction = Nothing
	
	End Function

	Public Function FormatAndSaveXMLFileGMD(ByVal strOutputPath, ByVal strSaveFileInProgress, ByVal strSaveFile, ByRef xmlElementTables, _
												ByRef xmlElementPassedParams, ByRef xmlElementProcessingInfo)
	'*****************************************************************************************************************************************
	'*  Purpose:				Setup all database recordsets for processing
	'*  Arguments supplied:		None
	'*  Return Value:			0 to indicate success
	'*  Called by:				Mainline
	'*  Calls:					CreateAppendNewElement, CreateNewElement, CreateSetAndLinkAttribute, AppendChild
	'*  Requirements:			None
	'*****************************************************************************************************************************************
		Dim objFSO, xmlDocLocal, xmlElementRemoteData, xmlProcessingInstruction

		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set xmlDocLocal = CreateObject("Microsoft.XMLDOM")
		Call CreateAppendNewElement(xmlDocLocal, xmlDocLocal, xmlElementRemoteData, "RemoteData")
		'
		' Append PassedParams, ProcessingInfo, and Tables to the RemoteData Element
		'
		Call g_objXMLProcessing.AppendChild(xmlElementPassedParams, xmlElementRemoteData)
		Call g_objXMLProcessing.AppendChild(xmlElementProcessingInfo, xmlElementRemoteData)
		Call g_objXMLProcessing.AppendChild(xmlElementTables, xmlElementRemoteData)
		'
		' Format the XML file
		'
		Set xmlProcessingInstruction = xmlDocLocal.createProcessingInstruction("xml","version='1.0' encoding='utf-8'")
		xmlDocLocal.insertBefore xmlProcessingInstruction, xmlDocLocal.childNodes(0)
		'
		' Save the XML file
		'
		xmlDocLocal.save strOutputPath & strSaveFileInProgress
		'
		' Rename the file now that the write has completed
		'
		objFSO.MoveFile strOutputPath & strSaveFileInProgress, strOutputPath & strSaveFile
		'
		' Cleanup
		'
		Set objFSO = Nothing
		Set xmlDocLocal = Nothing
		Set xmlElementRemoteData = Nothing
		Set xmlProcessingInstruction = Nothing

	End Function

End Class

Class ClsBrowse
'==========================================================================
' ClsBrowse Class
'
' This VB class allows users on multiple Windows platforms (W2K - 2003 tested)
' to find and select a file (ChooseFile) or folder (ChooseFolder) through a GUI.
'
' Example of use:
'
' Option Explicit


' Dim objBrowse, strRetVal
' 
' Set objBrowse = new ClsBrowse
'
' strRetVal = objBrowse.ChooseFile( "File open title goes here" )
' strRetVal = objBrowse.ChooseFolder( "Choose File Location" )
' 
' WScript.Echo strRetVal
'
' Set objBrowse = Nothing
' WScript.Quit
'
'==========================================================================
	Private m_objIE, m_objFSO, m_objShell, m_blnWordInstalled, m_objWord

          
	Private Sub Class_Initialize()
	'*************************************************************************************************************************************
	'*  Purpose:				Construct the class
	'*  Arguments supplied:		None
	'*  Return Value:			None
	'*  Called by:				Operating system when object is instantiated
	'*  Calls:					None
	'*  Requirements:			None
	'*************************************************************************************************************************************
		Set m_objFSO = CreateObject("Scripting.FileSystemObject")
		Set m_objShell = CreateObject("Shell.Application")
		
 	End Sub

	Private Sub Class_Terminate()
	'*************************************************************************************************************************************
	'*  Purpose:				Destruct the class
	'*  Arguments supplied:		None
	'*  Return Value:			None
	'*  Called by:				Operating system when object is destroyed
	'*  Calls:					None
	'*  Requirements:			None
	'*************************************************************************************************************************************
		Set m_objFSO = Nothing 
		Set m_objShell = Nothing
		Set m_objIE = Nothing
		Set m_objWord = Nothing
		
	 End Sub

	Public Function CheckWordInstallation()
	'*************************************************************************************************************************************
	'*  Purpose:				Check to see if Excel is installed on this computer
	'*  Arguments supplied:		Look up
	'*  Return Value:			True if installed; False if not installed
	'*  Called by:				Main()
	'*  Calls:					None
	'*  Requirements:			Registry Constants
	'*************************************************************************************************************************************
		Dim intError
		
		On Error Resume Next
		Set m_objWord = CreateObject("Word.Application")
		intError = Err.Number
		On Error GoTo 0
		If (intError = 0) Then
			CheckWordInstallation = True
			m_objWord.WindowState = 2
			m_objWord.Visible = False
			m_objWord.DisplayAlerts = False
		Else
			CheckWordInstallation = False
		End If
	
	End Function

	Public Function ChooseFile( ByVal strTitle )
	'*************************************************************************************************************************************
	'*  Purpose:				Display a FileOpen box to the user in order to select a file
	'*  Arguments supplied:		None
	'*  Return Value:			None
	'*  Called by:				User
	'*  Calls:					None
	'*  Requirements:			None
	'*************************************************************************************************************************************
		Dim objFile, strDQ, blnISaySo, intCount, intErrNumber, strErrDescription, strRetVal, strFolderPath, strFileName
		Const msoFileDialogOpen = 1

		ChooseFile = ""
		m_blnWordInstalled = CheckWordInstallation()
		If (m_blnWordInstalled) Then
			m_objWord.ChangeFileOpenDirectory(CreateObject("Wscript.Shell").SpecialFolders("Desktop"))
			m_objWord.FileDialog(msoFileDialogOpen).Title = strTitle
			m_objWord.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
			If (m_objWord.FileDialog(msoFileDialogOpen).Show = -1) Then
				'
				' There will only be 1 file returned
				'
				For Each objFile In m_objWord.FileDialog(msoFileDialogOpen).SelectedItems
					ChooseFile = objFile
					m_objWord.Visible = False
					m_objWord.Quit
					Set m_objWord = Nothing
					Exit Function
				Next 
			Else
				WScript.Echo "You selected 'Cancel'...Processing complete."
				m_objWord.Visible = False
				m_objWord.Quit
				Set m_objWord = Nothing
				WScript.Quit
			End If
			m_objWord.Visible = False
			m_objWord.Quit
			Set m_objWord = Nothing
		End If
		'
		' Office/Word not installed - see if IE will work...
		'
		strDQ = Chr(34) ' Double Quotes
		blnISaySo = True
		intCount = 0
		While blnISaySo
			On Error Resume Next
			Set m_objIE = CreateObject("InternetExplorer.Application")
			intErrNumber = Err.Number
			strErrDescription = Err.Description
			On Error GoTo 0
			If (intErrNumber) Then
				intCount = intCount + 1
				If (intCount >= 10) Then
					WScript.Echo "InternetExplorer.Application could not be created...Program execution ceasing"
					WScript.Quit
				End If
				WScript.Sleep 1000
			Else
				blnISaySo = False
			End If
		Wend
		m_objIE.Visible = False
		m_objIE.Offline = True
		m_objIE.Navigate("about:blank")

		Do Until m_objIE.ReadyState = 4
		Loop
		
		m_objIE.Document.Write "<HTML><BODY><INPUT ID=" & strDQ & "Fil" & strDQ & "Type=" & strDQ & "file" & strDQ & "></BODY></HTML>"
		With m_objIE.Document.all.Fil
			.focus
			.click
			strRetVal = .value
		End With
		m_objIE.Quit
	    Set m_objIE = Nothing
		'
		' If the filepath is "...fakepath\..." it is because of IE settings and the call will fail.
		'
		' This added "just in case" because a web path is returned in some Windows versions.
		'
		strRetVal = Replace(strRetVal, "%20", " ")
		If (strRetVal = "") Then
			WScript.Echo "Cancel selected.  No further processing will occur...Too Bad...So sad...ByeBye."
			WScript.Quit
		End If
		If (InStr(1, strRetVal, "fakepath", vbTextCompare) > 0) Then
			strFileName = m_objFSO.GetFileName(strRetVal)
			MsgBox "IE issue determining filepath.  Please choose the filepath (Folder) where " & strFileName & " is located", vbOKOnly
			strFolderPath = ChooseFolder("Choose the folder where " & strFileName & " is located")
			If ( strFolderPath = "" ) Then
				WScript.Echo "Invalid file location specified.  Program abending."
				WScript.Sleep 1000
				WScript.Quit
			End If
			'
			' The strDLLPath should contain the valid location of the .DLL files (without a '\')
			'				
			strFolderPath = strFolderPath & "\"
			strRetVal = strFolderPath & strFileName
			If (m_objFSO.FileExists(strRetVal)) Then
				ChooseFile = strRetVal
			Else
				WScript.Echo "File name selected doesn't exist in current path...Please try again."
				WScript.Quit
			End If
			Exit Function
		End If
		If (InStr(strRetVal, ":") = 0) Then
			WScript.Echo "Selected file only contains file name (file path is missing due to IE settings)...Please try again."
			WScript.Quit
		End If
		If (m_objFSO.FileExists(strRetVal)) Then
			ChooseFile = strRetVal
		End If
		
	End Function

	Function ChooseFolder( ByVal strCaption )
	'*************************************************************************************************************************************
	'*  Purpose:				Allows the user to browse for a folder using the shell.application object.
	'*  Arguments supplied:		strScriptPath and strFileType are passed so this function is generic.
	'*  Return Value:			Path that the user selected (may be blank if cancel was selected).
	'*  Called by:				InitializeVariables
	'*  Calls:					None
	'*  Requirements:			None
	'*************************************************************************************************************************************
		'
		' This is a version that shows files and will Not return Namespaces.
		'
		' Syntax: objShell.BrowseForFolder( Handle, Title, Options, RootFolder )
		'
		' Return Type: String
		'
		' Parameters:
		'
		' 	Handle: This is the numeric value of the owner application.  When calling this from a script, you can specify a 0 as the handle ID.
		'
		' 	Title: The text to go at the top of the browse dialog box, but not in the title bar area.
		'
		'	Options: An option constant or combination of contstants where the constants are combined using the OR.
		'
		'	RootFolder: The folder to start the browsing in. This can either be a string literal, or a constant. 
		'
		Dim objFolder, intBrowseInfo, blnISaySo, objShell, strTitle, intRetVal, strPath, intPOS
		Dim strFolderName, strParentName, objParentFolder, strDesk

		Const BIF_RETURNONLYFSDIRS = &H1		' Only return file system directories. If the user selects folders that are not part of the
												' file system, the OK button is grayed.
		Const BIF_DONTGOBELOWDOMAIN = &H2		' Do not include network folders below the domain level in the dialog box's tree view control.
		Const BIF_STATUSTEXT = &H4				' Include a status area in the dialog box. The callback function can set the status text by
												' sending messages to the dialog box. This flag is not supported when BIF_NEWDIALOGSTYLE Is
												' specified.
		Const BIF_RETURNFSANCESTORS = &H8		' Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder
												' in the namespace hierarchy. If the user selects an ancestor of the root folder that is not part
												' of the file system, the OK button is grayed.
		Const BIF_EDITBOX = &H10				' Version 4.71. Include an edit control in the browse dialog box that allows the user to type
												' the name of an item.
		Const BIF_VALIDATE = &H20				' Version 4.71. If the user types an invalid name into the edit box, the browse dialog box will
												' call the application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message. This flag Is
												' ignored if BIF_EDITBOX is not specified.
		Const BIF_NEWDIALOGSTYLE = &H30			' Version 5.0. Use the new user interface. Setting this flag provides the user with a larger
												' dialog box that can be resized. The dialog box has several new capabilities including: 
												' drag-and-drop capability within the dialog box, reordering, shortcut menus, new folders, 
												' delete, and other shortcut menu commands. To use this flag, you must call OleInitialize or 
												' CoInitialize before calling SHBrowseForFolder.
		Const BIF_USENEWUI = &H40				' Version 5.0. Use the new user interface, including an edit box. This flag is equivalent To
												' BIF_EDITBOX | BIF_NEWDIALOGSTYLE. To use BIF_USENEWUI, you must call OleInitialize or 
												' CoInitialize before calling SHBrowseForFolder.
		Const BIF_BROWSEINCLUDEURLS = &H80		' Version 5.0. The browse dialog box can display URLs. The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES 
												' flags must also be set. If these three flags are not set, the browser dialog box will reject
												' URLs. Even when these flags are set, the browse dialog box will only display URLs if the
												' folder that contains the selected item supports them. When the folder's 
												' IShellFolder::GetAttributesOf method is called to request the selected item's attributes, the
												' folder must set the SFGAO_FOLDER attribute flag. Otherwise, the browse dialog box will not 
												' display the URL.
		Const BIF_UAHINT = &H100				' Version 6.0. When combined with BIF_NEWDIALOGSTYLE, adds a usage hint to the dialog box in place
												' of the edit box. BIF_EDITBOX overrides this flag.
		Const BIF_NONEWFOLDER = &H200			' Version 6.0. Do not include the New Folder button in the browse dialog box.
		Const BIF_NOTRANSLATETARGETS = &H400	' Version 6.0. When the selected item is a shortcut, return the PIDL of the shortcut itself
												' rather than its target.
		Const BIF_BROWSEFORCOMPUTER = &H1000	' Only return computers. If the user selects anything other than a computer, the OK button is grayed.
		Const BIF_BROWSEFORPRINTER = &H2000		' Only allow the selection of printers.  If the user selects anything other than a printer, the OK
												' button is grayed.  In Microsoft Windows XP, the best practice is to use an XP-style dialog, 
												' setting the root of the dialog to the Printers and Faxes folder (CSIDL_PRINTERS).
		Const BIF_BROWSEINCLUDEFILES = &H4000	' Version 4.71. The browse dialog box will display files as well as folders.
		Const BIF_SHAREABLE = &H8000			' Version 5.0. The browse dialog box can display shareable resources on remote systems. It Is
												' intended for applications that want to expose remote shares on a local system. The 
												' BIF_NEWDIALOGSTYLE flag must also be set.
		'
		' Folder constants
		'
		'Public Enum efbrCSIDLConstants
		Const CSIDL_DESKTOP = &H0					'(desktop)
		Const CSIDL_INTERNET = &H1					'Internet Explorer (icon on desktop)
		Const CSIDL_PROGRAMS = &H2					'Start Menu\Programs
		Const CSIDL_CONTROLS = &H3					'My Computer\Control Panel
		Const CSIDL_PRINTERS = &H4					'My Computer\Printers
		Const CSIDL_PERSONAL = &H5					'My Documents
		Const CSIDL_FAVORITES = &H6					'(user name)\Favorites
		Const CSIDL_STARTUP = &H7					'Start Menu\Programs\Startup
		Const CSIDL_RECENT = &H8					'(user name)\Recent
		Const CSIDL_SENDTO = &H9					'(user name)\SendTo
		Const CSIDL_BITBUCKET = &HA					'(desktop)\Recycle Bin
		Const CSIDL_STARTMENU = &HB					'(user name)\Start Menu
		Const CSIDL_DESKTOPDIRECTORY = &H10			'(user name)\Desktop
		Const CSIDL_DRIVES = &H11					'My Computer
		Const CSIDL_NETWORK = &H12					'Network Neighborhood
		Const CSIDL_NETHOOD = &H13					'(user name)\nethood
		Const CSIDL_FONTS = &H14					'windows\fonts
		Const CSIDL_TEMPLATES = &H15			
		Const CSIDL_COMMON_STARTMENU = &H16			'All Users\Start Menu
		Const CSIDL_COMMON_PROGRAMS = &H17			'All Users\Programs
		Const CSIDL_COMMON_STARTUP = &H18			'All Users\Startup
		Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19	'All Users\Desktop
		Const CSIDL_APPDATA = &H1A					'(user name)\Application Data
		Const CSIDL_PRINTHOOD = &H1B				'(user name)\PrintHood
		Const CSIDL_LOCAL_APPDATA = &H1C			'(user name)\Local Settings\Application Data (non roaming)
		Const CSIDL_ALTSTARTUP = &H1D				'non localized startup
		Const CSIDL_COMMON_ALTSTARTUP = &H1E		'non localized common startup
		Const CSIDL_COMMON_FAVORITES = &H1F			
		Const CSIDL_INTERNET_CACHE = &H20			
		Const CSIDL_COOKIES = &H21					
		Const CSIDL_HISTORY = &H22					
		Const CSIDL_COMMON_APPDATA = &H23			'All Users\Application Data
		Const CSIDL_WINDOWS = &H24					'GetWindowsDirectory()
		Const CSIDL_SYSTEM = &H25					'GetSystemDirectory()
		Const CSIDL_PROGRAM_FILES = &H26			'C:\Program Files
		Const CSIDL_MYPICTURES = &H27				'C:\Program Files\My Pictures
		Const CSIDL_PROFILE = &H28					'USERPROFILE
		Const CSIDL_PROGRAM_FILES_COMMON = &H2B		'C:\Program Files\Common
		Const CSIDL_COMMON_TEMPLATES = &H2D			'All Users\Templates
		Const CSIDL_COMMON_DOCUMENTS = &H2E			'All Users\Documents
		Const CSIDL_COMMON_ADMINTOOLS = &H2F		'All Users\Start Menu\Programs\Administrative Tools
		Const CSIDL_ADMINTOOLS = &H30				'(user name)\Start Menu\Programs\Administrative Tools
		Const CSIDL_FLAG_CREATE = &H8000			'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
		Const CSIDL_FLAG_DONT_VERIFY = &H4000		'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
		Const CSIDL_FLAG_MASK = &HFF00				'mask for all possible flag values
		'End Enum

		intBrowseInfo = BIF_BROWSEINCLUDEFILES
'		intBrowseInfo = BIF_BROWSEINCLUDEFILES + BIF_EDITBOX + BIF_BROWSEINCLUDEFILES + BIF_VALIDATE
'		strPrompt = "Choose File location"
		Set objShell = CreateObject("Shell.Application")
		strTitle = "Error - Invalid Folder Selection"

'		intRetVal = MsgBox( "One or more of the " & strFileType & " files wasn't found in the " & strScriptPath & " directory.  Would you like to select the location?" & _
'							VbCrLf & VbCrLf & "Note: If you select cancel the script will abend as these are required files", vbOKCancel Or vbDefaultButton1, "Process Settings" )
'		If ( intRetVal <> vbOK ) Then
'			WScript.Quit
'		End If
	
		blnISaySo = True
		Do While ( blnISaySo )
			On Error Resume Next
			Set objFolder = objShell.BrowseForFolder( &H0, strCaption, intBrowseInfo, CSIDL_DRIVES )
'			Set objFolder = objShell.BrowseForFolder( &H0, strCaption, intBrowseInfo, CSIDL_DESKTOP )
			If ( Err.Number = -2147024894 ) Then
				'
				' The user selected a file instead of a folder.  Give them a message box.
				'
				intRetVal = MsgBox( "ERROR: You selected a file.  Please select a FOLDER where the files are located." & VbCrLf & VbCrLf & _
									"Note: Files within folders are only displayed to aid in locating the correct folder" ,, strTitle )
			Else
				blnISaySo = False
			End If
			On Error goto 0
		Loop
		On Error goto 0
		'
		' The user either selected something or pressed cancel
		'		
		If ( objFolder Is Nothing ) Then
			'
			' The objFolder Is Nothing is True (didn't get assigned a value).  This means
			' the user selected "Cancel".  Return nothing to the user.  The calling routine
			' should probably call quit if no valid path is returned.
			'
			ChooseFolder = ""
			Exit Function
		End If
		'
		' The "objFolder Is Nothing" is False (user selected something).
		'
		strFolderName = objFolder.Title
	
		strParentName = "a"
		Do While strParentName <> ""
			On Error Resume Next
			Set objParentFolder = objFolder.ParentFolder
			strParentName = objParentFolder.Title
			If ( Err.Number <> 0 ) Then
				'
				' An error here means no parent folder and no : has been found below
				' so it must be a drive or namespace (control panel, etc.)
				'
				intPOS = InStr( strFolderName, ":" ) 
				If ( intPOS = 0 ) Then                   '--it's a namespace or namespace path. check For Desktop.
					If ( Left( strFolderName, 6 ) = "Deskto" ) Then
						Set objShell = CreateObject("WScript.Shell")
						strDesk = objShell.Specialfolders("Desktop")
						Set objShell = Nothing
						If ( Len( strFolderName ) = 7 ) Then
							strFolderName = strDesk
						Else
							strFolderName = Right( strFolderName, ( Len( strFolderName ) - 7 ))
							strFolderName = strDesk & strFolderName
						End If
						'
						' remove %20 just in case.
						'
						ChooseFolder = Replace( strFolderName, "%20", " " )
						Set objShell = Nothing
					Else
						ChooseFolder = ""
					End If
					Set objFolder = Nothing
					Set objParentFolder = Nothing   
					On Error goto 0
					Exit Function
				Else                                '--it's a drive. extract root folder path (ex.: C:\ )
					strParentName = Mid( strFolderName, ( intPOS - 1 ), 2 )
					ChooseFolder = strParentName & "\"
					Set objFolder = Nothing
					Set objParentFolder = Nothing
					On Error goto 0
					Exit Function
				End If
			End If
       
			If ( Len( strParentName ) > 0 ) Then   '--look For a colon. If found Then quit Loop. If Not Then keep going.
				intPOS = InStr( strParentName, ":" )
				If ( intPOS = 0 ) Then
					'
					' No colon - add folder name to path and keep going.
					'
					strFolderName = strParentName & ( "\" & strFolderName )
				Else
					'
					' Colon found - get root folder, add to path and quit Loop.
					'
					strParentName = Mid( strParentName, ( intPOS - 1 ), 2 )
					strFolderName = strParentName & ( "\" & strFolderName )
					Exit Do
				End If
			End If
            '
			' If we are still processing then the path hasn't been found.
			' Set the parent folder as objFolder object and redo the Loop
			'
			Set objFolder = objParentFolder
			On Error goto 0
		Loop  
		On Error goto 0
		Set objFolder = Nothing
		Set objParentFolder = Nothing
		Set objShell = Nothing
		'
		' remove %20 just in case.
		'
		ChooseFolder = Replace( strFolderName, "%20", " " )

	End Function
    
End Class



'
' Functions
'
Function LogIt(ByVal xmlElement, ByRef objLogAndTrace)
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

Function ValidateEntryExists(ByRef rsEntries, ByVal strFilterValue, ByRef valData)
'*****************************************************************************************************************************************
'*  Purpose:				Checks the rsEntries recordset for a specific value and returns associated data.
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				BuildAddRemovePrograms
'*  Calls:					None
'*  Requirements:			None
'*****************************************************************************************************************************************

	If (Not rsEntries.BOF) Then
		rsEntries.MoveFirst
	End If
	rsEntries.Filter = "Entry = '" & strFilterValue & "'"
	If (rsEntries.RecordCount > 0) Then
		ValidateEntryExists = True
		valData = rsEntries("Data")
	Else
		ValidateEntryExists	= False
	End If
	rsEntries.Filter = 0

End Function

Function ProcessXMLFile(ByVal strFileName, ByVal blnProcessXML_SCCM, ByVal blnProcessXML_HBSS, ByRef objSCCMValidation, _
							ByRef objHBSSValidation, ByVal blnVerbose)
'*****************************************************************************************************************************************
'*  Purpose:				Processes output XML file for a specific computer
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				ProcessXMLFiles
'*  Calls:					BuildPath
'*  Requirements:			Generic Constants (library)
'*****************************************************************************************************************************************
	Dim strFileToProcess, xmlDoc, colRowNodes, objRowNode, colAttributes, objAttribute, strAttributeName, valAttributeValue
	Dim strComputerFQDN, strComputerName, strIPAddress, strOSName, strOSBuildNumber, dtLastBootupTime, strOSBuildType, intOSType
	Dim strOSType, intProductType, strProductType, strOSVersion, intAddressWidth, intPercentFreeC, strManuallyAssignedSiteCode
	Dim strGPOAssignedSiteCode, strEncryptionSubject, blnEncryptionSubjectMatch, strEncryptionSubjectMatch, dtEncryptionNotBefore
	Dim dtEncryptionNotAfter, 	strSigningSubject, blnSigningSubjectMatch, strSigningSubjectMatch, dtSigningNotBefore, dtSigningNotAfter
	Dim strWUAState, strWUAUStartMode, strSCCMExecState, strSCCMExecStartMode, strBITSState, strBITSStartMode, strWinMgmtState
	Dim strWinMgmtStartMode, strGPClientState, strGPClientStartMode, strMcFrameworkState, strMcFrameworkStartMode, strMcShieldState
	Dim strMcShieldStartMode, strSCCMVersion, blnWMIGoodToGo, strWMIGoodToGo, blnWMIRegGoodToGo, strWMIRegGoodToGo, blnRemoteRegGoodToGo
	Dim strRemoteRegGoodToGo, strWindowsUpdateServer, strNameServers, dtAVDatDate, dtCatalogVersionDate, strSiteListName, strSiteListIP
	Dim strSiteListPort, strEPORegistryName, strEPORegistryIP, strEPORegistryPort, strAgentGUID, strLastLoggedOnUser, blnPendingReboot
	Dim strPendingReboot, strSMSUID, strPreviousSMSUID, dtLastChangedSMSUID, strClientSite, strCurrentMP, strADSiteName, strSDCVersion
	Dim strWUAVersion, strCCMSetupState, strCCMSetupStartMode, strLanmanServerState, strLanmanServerStartMode, strRPCSSState
	Dim strRPCSSStartMode, strSMSTSMGRState, strSMSTSMGRStartMode, strRemoteRegistryState, strRemoteRegistryStartMode, blnEnableDCOM
	Dim strEnableDCOM, intRunningAdvertisements, blnSCCMHealthy, strMostRecentSMSFolder, strMostRecentSMSFolderDate, blnHBSSHealthy
	Dim dtLastASCTime, dtPropsVersionDate, strAgentWakeUpPort, strMAVersion, blnFWEnabled, strFWEnabled, strVSEVersion, strHIPSVersion
	Dim strDLPVersion, dtCreationTimestamp, intDaysSinceLastBootup, strSCCMClientInstallation, strData

	'
	' Build full path for FileToProcess
	'
	strFileToProcess = g_strXMLOutputPath & strFileName

	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	xmlDoc.async = False
	xmlDoc.load(strFileToProcess)		' Load the entire document before continuing processing
	xmlDoc.validateOnParse = False
	xmlDoc.preserveWhiteSpace = True
	xmlDoc.resolveExternals = False

' 	if (g_xmlDoc.parseError.errorCode != 0) {
' 	   var myErr = xmlDoc.parseError;
' 	   WScript.Echo("You have error " + myErr.reason);
' 	} else {
' 	   myErr = xmlDoc.parseError;
' 	   if (myErr.errorCode != 0) {
' 	      WScript.Echo("You have error " + myErr.reason);
' 	   }
' 	}
'	g_xmlDoc.load strFileToProcess

	Set colRowNodes = xmlDoc.documentElement.selectNodes("//Tables/Table/Row")
	For Each objRowNode In colRowNodes
		Set colAttributes = objRowNode.Attributes
		For Each objAttribute In colAttributes
			strAttributeName = objAttribute.Name
			valAttributeValue = objAttribute.Value
			Select Case UCase(strAttributeName)
				Case "CLHCOMPUTERFQDN"
					strComputerFQDN = valAttributeValue
				Case "CLHCOMPUTERNAME"
					strComputerName = valAttributeValue
				Case "CLHIPADDRESS"
					strIPAddress = valAttributeValue
				Case "CLHOSNAME"
					strOSName = valAttributeValue
				Case "CLHOSBUILDNUMBER"
					strOSBuildNumber = valAttributeValue
				Case "CLHLASTBOOTUPTIME"
					If (CDate(valAttributeValue) = CDate(DEFAULT_DATE)) Then
						dtLastBootupTime = "NOT AVAILABLE"
						intDaysSinceLastBootup = -1
					Else
						dtLastBootupTime = valAttributeValue
						intDaysSinceLastBootup = DateDiff("d", dtLastBootupTime, Now)
					End If
				Case "CLHOSBUILDTYPE"
					strOSBuildType = valAttributeValue
				Case "CLHOSTYPE"
					intOSType = valAttributeValue
				Case "CLHOSTYPETEXT"
					strOSType = valAttributeValue
				Case "CLHPRODUCTTYPE"
					intProductType = valAttributeValue
				Case "CLHPRODUCTTYPETEXT"
					strProductType = valAttributeValue
				Case "CLHOSVERSION"
					strOSVersion = valAttributeValue
				Case "CLHADDRESSWIDTH"
					intAddressWidth = valAttributeValue
				Case "CLHPERCENTFREEC"
					intPercentFreeC = valAttributeValue
				Case "CLHMANUALLYASSIGNEDSITECODE"
					strManuallyAssignedSiteCode = valAttributeValue
				Case "CLHGPOASSIGNEDSITECODE"
					strGPOAssignedSiteCode = valAttributeValue
				Case "CLHENCRYPTIONSUBJECT"
					strEncryptionSubject = valAttributeValue
				Case "CLHENCRYPTIONSUBJECTMATCH"
					blnEncryptionSubjectMatch = valAttributeValue
					If (blnEncryptionSubjectMatch) Then
						strEncryptionSubjectMatch = "GOOD"
					Else
						strEncryptionSubjectMatch = "BAD"
					End If
				Case "CLHENCRYPTIONNOTBEFORE"
					If (CDate(valAttributeValue) = CDate(DEFAULT_DATE)) Then
						dtEncryptionNotBefore = "NOT AVAILABLE"
					Else
						dtEncryptionNotBefore = valAttributeValue
					End If
				Case "CLHENCRYPTIONNOTAFTER"
					If (CDate(valAttributeValue) = CDate(DEFAULT_DATE)) Then
						dtEncryptionNotAfter = "NOT AVAILABLE"
					Else
						dtEncryptionNotAfter = valAttributeValue
					End If
				Case "CLHSIGNINGSUBJECT"
					strSigningSubject = valAttributeValue
				Case "CLHSIGNINGSUBJECTMATCH"
					blnSigningSubjectMatch = valAttributeValue
					If (blnSigningSubjectMatch) Then
						strSigningSubjectMatch = "GOOD"
					Else
						strSigningSubjectMatch = "BAD"
					End If
				Case "CLHSIGNINGNOTBEFORE"
					If (CDate(valAttributeValue) = CDate(DEFAULT_DATE)) Then
						dtSigningNotBefore = "NOT AVAILABLE"
					Else
						dtSigningNotBefore = valAttributeValue
					End If
				Case "CLHSIGNINGNOTAFTER"
					If (CDate(valAttributeValue) = CDate(DEFAULT_DATE)) Then
						dtSigningNotAfter = "NOT AVAILABLE"
					Else
						dtSigningNotAfter = valAttributeValue
					End If
				Case "CLHWUAUSTATE"
					strWUAState = valAttributeValue
				Case "CLHWUAUSTARTMODE"
					strWUAUStartMode = valAttributeValue
				Case "CLHCCMEXECSTATE"
					strSCCMExecState = valAttributeValue
				Case "CLHCCMEXECSTARTMODE"
					strSCCMExecStartMode = valAttributeValue
				Case "CLHBITSSTATE"
					strBITSState = valAttributeValue
				Case "CLHBITSSTARTMODE"
					strBITSStartMode = valAttributeValue
				Case "CLHWINMGMTSTATE"
					strWinMgmtState = valAttributeValue
				Case "CLHWINMGMTSTARTMODE"
					strWinMgmtStartMode = valAttributeValue
				Case "CLHGPCLIENTSTATE"
					strGPClientState = valAttributeValue
				Case "CLHGPCLIENTSTARTMODE"
					strGPClientStartMode = valAttributeValue
				Case "CLHCCMSETUPSTATE"
					strCCMSetupState = valAttributeValue
				Case "CLHCCMSETUPSTARTMODE"
					strCCMSetupStartMode = valAttributeValue
				Case "CLHLANMANSERVERSTATE"
					strLanmanServerState = valAttributeValue
				Case "CLHLANMANSERVERSTARTMODE"
					strLanmanServerStartMode = valAttributeValue
				Case "CLHREMOTEPROCEDURECALLSTATE"
					strRPCSSState = valAttributeValue
				Case "CLHREMOTEPROCEDURECALLSTARTMODE"
					strRPCSSStartMode = valAttributeValue
				Case "CLHSMSTASKSEQUENCEAGENTSTATE"
					strSMSTSMGRState = valAttributeValue
				Case "CLHSMSTASKSEQUENCEAGENTSTARTMODE"
					strSMSTSMGRStartMode = valAttributeValue
				Case "CLHREMOTEREGISTRYSTATE"
					strRemoteRegistryState = valAttributeValue
				Case "CLHREMOTEREGISTRYSTARTMODE"
					strRemoteRegistryStartMode = valAttributeValue
				Case "CLHSCCMVERSION"
					strSCCMVersion = valAttributeValue
				Case "CLHWMIGOODTOGO"
					blnWMIGoodToGo = valAttributeValue
					If (blnWMIGoodToGo) Then
						strWMIGoodToGo = "YES"
					Else
						strWMIGoodToGo = "NO"
					End If
				Case "CLHWMIREGGOODTOGO"
					blnWMIRegGoodToGo = valAttributeValue
					If (blnWMIRegGoodToGo) Then
						strWMIRegGoodToGo = "YES"
					Else
						strWMIRegGoodToGo = "NO"
					End If
				Case "CLHREMOTEREGGOODTOGO"
					blnRemoteRegGoodToGo = valAttributeValue
					If (blnRemoteRegGoodToGo) Then
						strRemoteRegGoodToGo = "YES"
					Else
						strRemoteRegGoodToGo = "NO"
					End If
				Case "CLHWINDOWSUPDATESERVER"
					strWindowsUpdateServer = valAttributeValue
				Case "CLHNAMESERVERS"
					strNameServers = valAttributeValue
				Case "CLHLASTLOGGEDONUSER"
					strLastLoggedOnUser = valAttributeValue
				Case "CLHPENDINGREBOOT"
					blnPendingReboot = valAttributeValue
					If (blnPendingReboot) Then
						strPendingReboot = "YES"
					Else
						strPendingReboot = "NO"
					End If
				Case "CLHSMSUID"
					strSMSUID = valAttributeValue
				Case "CLHPREVIOUSSMSUID"
					strPreviousSMSUID = valAttributeValue
				Case "CLHLASTCHANGEDSMSUID"
					If (CDate(valAttributeValue) = CDate(DEFAULT_DATE)) Then
						dtLastChangedSMSUID = "NOT AVAILABLE"
					Else
						dtLastChangedSMSUID = valAttributeValue
					End If
				Case "CLHCLIENTSITE"
					strClientSite = valAttributeValue
				Case "CLHCURRENTMP"
					strCurrentMP = valAttributeValue
				Case "CLHADSITENAME"
					strADSiteName = valAttributeValue
				Case "CLHSDCVERSION"
					strSDCVersion = valAttributeValue
				Case "CLHWUAVERSION"
					strWUAVersion = valAttributeValue
				Case "CLHRUNNINGADVERTISEMENTS"
					intRunningAdvertisements = valAttributeValue
				Case "CLHMOSTRECENTSMSFOLDER"
					strMostRecentSMSFolder = valAttributeValue
				Case "CLHMOSTRECENTSMSFOLDERDATE"
					'
					' MostRecentSMSUpdate
					'
					If (CDate(valAttributeValue) = CDate(DEFAULT_DATE)) Then
						strMostRecentSMSFolderDate = "NOT AVAILABLE"
					Else
						strMostRecentSMSFolderDate = valAttributeValue
					End If
				Case "CLHENABLEDCOM"
					blnEnableDCOM = valAttributeValue
					If (blnEnableDCOM) Then
						blnEnableDCOM = "YES"
					Else
						blnEnableDCOM = "NO"
					End If
				Case "CLHSCCMHEALTHY"
					blnSCCMHealthy = valAttributeValue
				'
				' HBSS settings
				'
				Case "CLHAVDATDATE"
					If (CDate(valAttributeValue) = CDate(DEFAULT_DATE)) Then
						dtAVDatDate = "NOT AVAILABLE"
					Else
						dtAVDatDate = valAttributeValue
					End If
				Case "CLHCATALOGVERSIONDATE"
					If (CDate(valAttributeValue) = CDate(DEFAULT_DATE)) Then
						dtCatalogVersionDate = "NOT AVAILABLE"
					Else
						dtCatalogVersionDate = valAttributeValue
					End If
				Case "CLHSITELISTNAME"
					strSiteListName = valAttributeValue
				Case "CLHSITELISTIP"
					strSiteListIP = valAttributeValue
				Case "CLHSITELISTPORT"
					strSiteListPort = valAttributeValue
				Case "CLHEPOREGISTRYNAME"
					strEPORegistryName = valAttributeValue
				Case "CLHEPOREGISTRYIP"
					strEPORegistryIP = valAttributeValue
				Case "CLHEPOREGISTRYPORT"
					strEPORegistryPort = valAttributeValue
				Case "CLHAGENTGUID"
					strAgentGUID = valAttributeValue
				Case "CLHMCFRAMEWORKSTATE"
					strMcFrameworkState = valAttributeValue
				Case "CLHMCFRAMEWORKSTARTMODE"
					strMcFrameworkStartMode = valAttributeValue
				Case "CLHMCSHIELDSTATE"
					strMcShieldState = valAttributeValue
				Case "CLHMCSHIELDSTARTMODE"
					strMcShieldStartMode = valAttributeValue
				Case "CLHLASTASCTIME"
					If (CDate(valAttributeValue) = CDate(DEFAULT_DATE)) Then
						dtLastASCTime = "NOT AVAILABLE"
					Else
						dtLastASCTime = valAttributeValue
					End If
				Case "CLHPROPSVERSIONDATE"
					If (CDate(valAttributeValue) = CDate(DEFAULT_DATE)) Then
						dtPropsVersionDate = "NOT AVAILABLE"
					Else
						dtPropsVersionDate = valAttributeValue
					End If
				Case "CLHAGENTWAKEUPPORT"
					If (CInt(valAttributeValue) = -1) Then
						strAgentWakeUpPort = "NOT AVAILABLE"
					Else
						strAgentWakeUpPort = valAttributeValue
					End If
				Case "CLHMAVERSION"
					strMAVersion = valAttributeValue
				Case "CLHFWENABLED"
					blnFWEnabled = valAttributeValue
					If (blnFWEnabled) Then
						strFWEnabled = "YES"
					Else
						strFWEnabled = "NO"
					End If
				Case "CLHVSEVERSION"
					strVSEVersion = valAttributeValue
				Case "CLHHIPSVERSION"
					strHIPSVersion = valAttributeValue
				Case "CLHDLPVERSION"
					strDLPVersion = valAttributeValue
				Case "CLHHBSSHEALTHY"
					blnHBSSHealthy = valAttributeValue
				Case "CLHCREATIONTIMESTAMP"
					dtCreationTimestamp = valAttributeValue
				Case Else
			End Select
		Next
	Next
	If (((UCase(strCCMSetupState) = "UNKNOWN") Or (UCase(strCCMSetupState) = "NOT INSTALLED")) And (strSCCMVersion <> "")) Then
		strSCCMClientInstallation = "COMPLETE"
	Else
		strSCCMClientInstallation = "INCOMPLETE"
	End If
	If (blnProcessXML_SCCM) Then
		'
		' Write the data into the SCCM file
		'
		If (blnVerbose) Then
			strData = strComputerFQDN & vbTab & strComputerName & vbTab & strIPAddress & vbTab & strOSName & vbTab & strOSVersion & vbTab & _
						intAddressWidth & vbTab & strSDCVersion & vbTab & intPercentFreeC & vbTab & strLastLoggedOnUser & vbTab & strPendingReboot & vbTab & _
						dtLastBootupTime & vbTab & intDaysSinceLastBootup & vbTab & strMostRecentSMSFolderDate & vbTab & intRunningAdvertisements & vbTab & _
						strNameServers & vbTab & strManuallyAssignedSiteCode & vbTab & strGPOAssignedSiteCode & vbTab & strClientSite & vbTab & _
						strEncryptionSubject & vbTab & strEncryptionSubjectMatch & vbTab & dtEncryptionNotBefore & vbTab & dtEncryptionNotAfter & vbTab & _
						strSigningSubject & vbTab & strSigningSubjectMatch & vbTab & dtSigningNotBefore & vbTab & dtSigningNotAfter & vbTab & _
						strSMSUID & vbTab & strPreviousSMSUID & vbTab & dtLastChangedSMSUID & vbTab & strADSiteName & vbTab & strWUAState & vbTab & _
						strWUAUStartMode & vbTab & strSCCMExecState & vbTab & strSCCMExecStartMode & vbTab & strBITSState & vbTab & strBITSStartMode & vbTab & _
						strWinMgmtState & vbTab & strWinMgmtStartMode & vbTab & strWMIGoodToGo & vbTab & strWMIRegGoodToGo & vbTab & strRemoteRegistryState & vbTab & _
						strRemoteRegistryStartMode & vbTab & strRemoteRegGoodToGo & vbTab & blnEnableDCOM & vbTab & strGPClientState & vbTab & _
						strGPClientStartMode & vbTab & strLanmanServerState & vbTab & strLanmanServerStartMode & vbTab & strRPCSSState & vbTab & _
						strRPCSSStartMode & vbTab & strSMSTSMGRState & vbTab & strSMSTSMGRStartMode & vbTab & strCCMSetupState & vbTab & strCCMSetupStartMode & vbTab & _
						strSCCMClientInstallation & vbTab & strSCCMVersion & vbTab & strCurrentMP & vbTab & strWindowsUpdateServer & vbTab & strWUAVersion & vbTab & _
						dtCreationTimestamp 
		Else
			strData = strComputerFQDN & vbTab & strComputerName & vbTab & strIPAddress & vbTab & strSDCVersion & vbTab & intPercentFreeC & vbTab & _
						strPendingReboot & vbTab & intDaysSinceLastBootup & vbTab & strMostRecentSMSFolderDate & vbTab & strManuallyAssignedSiteCode & vbTab & _
						strGPOAssignedSiteCode & vbTab & strClientSite & vbTab & strEncryptionSubject & vbTab & strEncryptionSubjectMatch & vbTab & _
						strSigningSubject & vbTab & strSigningSubjectMatch & vbTab & strWUAState & vbTab & strWUAUStartMode & vbTab & strSCCMExecState & vbTab & _
						strSCCMExecStartMode & vbTab & strBITSState & vbTab & strBITSStartMode & vbTab & strWinMgmtState & vbTab & strWinMgmtStartMode & vbTab & _
						strWMIGoodToGo & vbTab & strWMIRegGoodToGo & vbTab & strRemoteRegistryState & vbTab & strRemoteRegistryStartMode & vbTab & _
						strRemoteRegGoodToGo & vbTab & strGPClientState & vbTab & strGPClientStartMode & vbTab & strSMSTSMGRState & vbTab & _
						strSMSTSMGRStartMode & vbTab & strSCCMClientInstallation & vbTab & strSCCMVersion & vbTab & strCurrentMP & vbTab & _
						strWindowsUpdateServer & vbTab & strWUAVersion & vbTab & dtCreationTimestamp 
		End If
		objSCCMValidation.WriteLine(strData)
	End If
	If (blnProcessXML_HBSS) Then
		'
		' Write the data into the HBSS file
		'
		strData = strComputerFQDN & vbTab & strComputerName & vbTab & strIPAddress & vbTab & strOSName & vbTab & strOSVersion & vbTab & intAddressWidth & vbTab & _
					strPendingReboot & vbTab & intDaysSinceLastBootup & vbTab & strRemoteRegistryState & vbTab & strRemoteRegistryStartMode & vbTab & _
					strRemoteRegGoodToGo & vbTab & strMcFrameworkState & vbTab & strMcFrameworkStartMode & vbTab & strMcShieldState & vbTab & _
					strMcShieldStartMode & vbTab & dtAVDatDate & vbTab & dtCatalogVersionDate & vbTab & strSiteListName & vbTab & strSiteListIP & vbTab & _
					strSiteListPort & vbTab & strEPORegistryName & vbTab & strEPORegistryIP & vbTab & strEPORegistryPort & vbTab & strAgentGUID & vbTab & _
					dtLastASCTime & vbTab & dtPropsVersionDate & vbTab & strMAVersion & vbTab & strVSEVersion & vbTab & strHIPSVersion & vbTab & _
					strDLPVersion & vbTab & strAgentWakeUpPort & vbTab & strFWEnabled & vbTab & dtCreationTimestamp 
		objHBSSValidation.WriteLine(strData)
	End If
	'
	' Cleanup
	'
	Set xmlDoc = Nothing
	Set colRowNodes = Nothing
	Set colAttributes = Nothing

End Function

Function ProcessXMLFiles(ByRef colFiles, ByVal blnProcessXML_SCCM, ByVal blnProcessXML_HBSS, ByVal blnVerbose)
'*****************************************************************************************************************************************
'*  Purpose:				Processes all XML output files
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				Mainline
'*  Calls:					BuildDateString, ProcessXMLFile
'*  Requirements:			ADO Constants
'*****************************************************************************************************************************************
	Dim rsFilesToProcess, objFile, strFileExtension, strFileName, strSCCMValidation, objSCCMValidation, strData, strHBSSValidation
	Dim objHBSSValidation

	'
	' Create Recordset
	'
	Set rsFilesToProcess = CreateObject("ADODB.Recordset")
	rsFilesToProcess.Fields.Append "FileName", adVarChar, 255
	rsFilesToProcess.Open

	For Each objFile In colFiles
		'
		' Get the file extension
		'
		strFileExtension = LCase(g_objFSO.GetExtensionName(objFile))
		'
		' Only process XML files
		'
		If (strFileExtension = "xml") Then
			strFileName = g_objFSO.GetFileName(objFile)
			rsFilesToProcess.AddNew
			rsFilesToProcess("FileName") = strFileName
			rsFilesToProcess.Update
		End If
	Next
	'
	' Make sure there are .xml records to process
	'
	If (rsFilesToProcess.RecordCount = 0) Then
		WScript.Echo "The XMLOutput folder contains no .xml files.  Please move XML files to folder and try again."
		WScript.Quit
	End If
	If (blnProcessXML_SCCM) Then
		'
		' Add the header information
		'
		If (blnVerbose) Then
			strSCCMValidation = g_strParentFolder & "Client_Health_Scan_SCCM_Results_Verbose_" & g_objFunctions.BuildDateString(Now) & ".tsv"
			strData = "ComputerFQDN" & vbTab & "ComputerName" & vbTab & "IPAddress" & vbTab & "OSName" & vbTab & "OSVersion" & vbTab & _
						"OSArchitecture" & vbTab & "SDCVersion" & vbTab & "PercentFreeC" & vbTab & "LastLoggedOnUser" & vbTab & "PendingReboot" & vbTab & _
						"LastBootupTime" & vbTab & "DaysSinceLastBootup" & vbTab & "MostRecentSMSUpdate" & vbTab & "RunningAdvertisements" & vbTab & _
						"DNSServers" & vbTab & "ManuallyAssignedSiteCode" & vbTab & "GPOAssignedSiteCode" & vbTab & "SMSActualSiteCode" & vbTab & _
						"EncryptionCertSubject" & vbTab & "EncryptionCertSubjectMatch" & vbTab & "EncryptionCertNotBefore" & vbTab & _
						"EncryptionCertNotAfter" & vbTab & "SigningCertSubject" & vbTab & "SigningCertSubjectMatch" & vbTab & "SigningCertNotBefore" & vbTab & _
						"SigningCertNotAfter" & vbTab & "SMSUID" & vbTab & "PreviousSMSUID" & vbTab & "LastChangedSMSUID" & vbTab & "ADSiteName" & vbTab & _
						"WindowsUpdateServiceState" & vbTab & "WindowsUpdateServiceStartMode" & vbTab & "SMSAgentHostServiceState" & vbTab & _
						"SMSAgentHostServiceStartMode" & vbTab & "BITSServiceState" & vbTab & "BITSServiceStartMode" & vbTab & "WinMgmtServiceState" & vbTab & _
						"WinMgmtServiceStartMode" & vbTab & "WMIGoodToGo" & vbTab & "WMIRegGoodToGo" & vbTab & "RemoteRegistryServiceState" & vbTab & _
						"RemoteRegistryServiceStartMode" & vbTab & "RemoteRegGoodToGo" & vbTab & "EnableDCOM" & vbTab & "GPClientServiceState" & vbtab & _
						"GPClientServiceStartMode" & vbTab & "LanmanServerServiceState" & vbTab & "LanmanServerServiceStartMode" & vbTab & _
						"RPCSSServiceState" & vbTab & "RPCSSServiceStartMode" & vbTab & "SMSTSMGRServiceState" & vbTab & "SMSTSMGRServiceStartMode" & vbTab & _
						"CCMSetupServiceState" & vbTab & "CCMSetupServiceStartMode" & vbTab & "SCCMClientInstallation" & vbTab & "SCCMVersion" & vbTab & _
						"CurrentManagementPoint" & vbTab & "WindowsUpdateServer(GPO)" & vbTab & "WindowsUpdateAgentVersion" & vbTab & "CreationTimestamp"
		Else
			strSCCMValidation = g_strParentFolder & "Client_Health_Scan_SCCM_Results_" & g_objFunctions.BuildDateString(Now) & ".tsv"
			strData = "ComputerFQDN" & vbTab & "ComputerName" & vbTab & "IPAddress" & vbTab & "SDCVersion" & vbTab & "PercentFreeC" & vbTab & _
						"PendingReboot" & vbTab & "DaysSinceLastBootup" & vbTab & "MostRecentSMSUpdate" & vbTab & "ManuallyAssignedSiteCode" & vbTab & _
						"GPOAssignedSiteCode" & vbTab & "SMSActualSiteCode" & vbTab & "EncryptionCertSubject" & vbTab & "EncryptionCertSubjectMatch" & vbTab & _
						"SigningCertSubject" & vbTab & "SigningCertSubjectMatch" & vbTab & "WindowsUpdateServiceState" & vbTab & _
						"WindowsUpdateServiceStartMode" & vbTab & "SMSAgentHostServiceState" & vbTab & "SMSAgentHostServiceStartMode" & vbTab & _
						"BITSServiceState" & vbTab & "BITSServiceStartMode" & vbTab & "WinMgmtServiceState" & vbTab & "WinMgmtServiceStartMode" & vbTab & _
						"WMIGoodToGo" & vbTab & "WMIRegGoodToGo" & vbTab & "RemoteRegistryServiceState" & vbTab & "RemoteRegistryServiceStartMode" & vbTab & _
						"RemoteRegGoodToGo" & vbTab & "GPClientServiceState" & vbtab & "GPClientServiceStartMode" & vbTab & "SMSTSMGRServiceState" & vbTab & _
						"SMSTSMGRServiceStartMode" & vbTab & "SCCMClientInstallation" & vbTab & "SCCMVersion" & vbTab & "CurrentManagementPoint" & vbTab & _
						"WindowsUpdateServer(GPO)" & vbTab & "WindowsUpdateAgentVersion" & vbTab & "CreationTimestamp"
		End If
		'
		' Create the .tsv file to store SMS data
		'
		Set objSCCMValidation = g_objFSO.OpenTextFile(strSCCMValidation, FOR_WRITE, CREATE_IF_NON_EXISTENT)
		objSCCMValidation.WriteLine(strData)
	End If
	If (blnProcessXML_HBSS) Then
		'
		' Add the header information
		'
		strData = "ComputerFQDN" & vbTab & "ComputerName" & vbTab & "IPAddress" & vbTab & "OSName" & vbTab & "OSVersion" & vbTab & "OSArchitecture" & vbTab & _
					"PendingReboot" & vbTab & "DaysSinceLastBootup" & vbTab & "RemoteRegistryServiceState" & vbTab & "RemoteRegistryServiceStartMode" & vbTab & _
					"RemoteRegGoodToGo" & vbTab & "McAfeeFrameworkServiceState" & vbTab & "McAfeeFrameworkServiceStartMode" & vbTab & _
					"McAfeeMcShieldServiceState" & vbTab & "McAfeeMcShieldServiceStartMode" & vbTab & "AVDatDate" & vbTab & _
					"McAfeeFrameworkCatalogVersionDate" & vbTab & "SiteListNames" & vbTab & "SiteListIPs" & vbTab & "SiteListPorts" & vbTab & _
					"EPORegistryNames" & vbTab & "EPORegistryIPs" & vbTab & "EPORegistryPorts" & vbTab & "AgentGUID" & vbTab & "LastASCTime" & vbTab & _
					"PropsVersionDate" & vbTab & "MAVersion" & vbTab & "VSEVersion" & vbTab & "HIPSVersion" & vbTab & "DLPVersion" & vbTab & _
					"AgentWakeUpPort" & vbTab & "FWEnabled" & vbTab & "CreationTimestamp"
		'
		' Create the .tsv file to store HBSS data
		'
		strHBSSValidation = g_strParentFolder & "Client_Health_Scan_HBSS_Results_" & g_objFunctions.BuildDateString(Now) & ".tsv"
		Set objHBSSValidation = g_objFSO.OpenTextFile(strHBSSValidation, FOR_WRITE, CREATE_IF_NON_EXISTENT)
		objHBSSValidation.WriteLine(strData)
	End If
	'
	' Process existing XML files
	'
	rsFilesToProcess.MoveFirst
	While Not rsFilesToProcess.EOF
		strFileName = rsFilesToProcess("FileName")
		Call ProcessXMLFile(strFileName, blnProcessXML_SCCM, blnProcessXML_HBSS, objSCCMValidation, objHBSSValidation, blnVerbose)
 		rsFilesToProcess.MoveNext
 	Wend
	'
	' Cleanup
	'
	Set rsFilesToProcess = Nothing
	Set objSCCMValidation = Nothing
	Set objHBSSValidation = Nothing

End Function

Function GetLastLoggedOnUser(ByRef objRemoteWMIServer, ByVal intFlag, ByRef strLastLoggedOnUser)
'*****************************************************************************************************************************************
'*  Purpose:				Gets the latest (or currently) logged on user.
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				ScanClientConfiguration
'*  Calls:					VerifyAndLoad
'*  Requirements:			ADO Constants
'*****************************************************************************************************************************************
	Dim colWMI, intErrNumber, strErrDescription, objWMI, dtLatestLogon, dtNewLastLogon

	On Error Resume Next
	Set colWMI = objRemoteWMIServer.ExecQuery("SELECT UserName FROM Win32_ComputerSystem",, intFlag)
	intErrNumber = Err.Number
	strErrDescription = Err.Description
	On Error GoTo 0
	If (intErrNumber=0) Then
		If (UCase(TypeName(colWMI)) = "SWBEMOBJECTSET") Then
			For Each objWMI In colWMI
				strLastLoggedOnUser = g_objFunctions.VerifyAndLoad(objWMI.UserName, vbString)
			Next
		End If
	End If

	If (strLastLoggedOnUser = "") Then
		dtLatestLogon = "00000000000000"
		On Error Resume Next
		Set colWMI = objRemoteWMIServer.ExecQuery("SELECT Name,LastLogon FROM Win32_NetworkLoginProfile",, intFlag) 
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
		If (intErrNumber=0) Then
			If (UCase(TypeName(colWMI)) = "SWBEMOBJECTSET") Then
				For Each objWMI in colWMI
					dtNewLastLogon = g_objFunctions.VerifyAndLoad(objWMI.LastLogon, vbString)
					If (dtNewLastLogon <> "") Then
						dtNewLastLogon = Left(dtNewLastLogon, 14)
						If (dtNewLastLogon > dtLatestLogon) Then
							dtLatestLogon = dtNewLastLogon
							strLastLoggedOnUser = g_objFunctions.VerifyAndLoad(objWMI.Name, vbString)
						End If
					End If
				Next
			End If
		End If
	End If

End Function

Function ParseSMSCertificateSubject(ByRef rsGeneric, ByRef strSubject)
'*****************************************************************************************************************************************
'*  Purpose:				Parses SMS Certificate information for Subject
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				ScanClientConfiguration
'*  Calls:					None
'*  Requirements:			None
'*****************************************************************************************************************************************
	Dim strLine, strCertificateSubject

	strCertificateSubject = ""
	If (Not rsGeneric.BOF) Then
		rsGeneric.MoveFirst
	End If
	While Not rsGeneric.EOF
		strLine = rsGeneric("SavedData")
		'
		' Look for this line -> Subject: CN=SMS, CN=52CSUDW3-410YW5
		'
		If (InStr(1, strLine, "Subject: ", vbTextCompare) > 0) Then
			strCertificateSubject = Trim(Split(strLine, ",", 2)(1))
			strSubject = Trim(Split(strCertificateSubject, "CN=", 2, vbTextCompare)(1))
		End If
		rsGeneric.MoveNext
	Wend

End Function

Function GetServiceSettings(ByRef objWMIServer, ByVal intFlag, ByVal blnWMIGoodToGo, ByVal strServiceName, ByVal strPassedParameter, _
								ByRef strState, ByRef blnStarted, ByRef strStartMode)
'*****************************************************************************************************************************************
'*  Purpose:				Gets the settings for the specified service.
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				GetClientConfiguration
'*  Calls:					ExecWMI, ExecCmdGeneric, DeleteAllRecordsetRows
'*  Requirements:			None
'*****************************************************************************************************************************************
	Dim strSQLQuery, intErrNumber, strErrDescription, colWMI, objWMI, strCommand, strLine

	'
	' See if service is running
	'
	strState = ""
	blnStarted = False
	strStartMode = ""

	If (blnWMIGoodToGo) Then
		strSQLQuery = "SELECT State,Started,StartMode FROM Win32_Service WHERE Name='" & strServiceName & "'"
		Call g_objFunctions.ExecWMI(objWMIServer, intErrNumber, strErrDescription, colWMI, strSQLQuery, intFlag, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
			If (colWMI.Count = 0) Then
				'
				' Service doesn't exist
				'
				strState = "Not Installed"
				blnStarted = False
				strStartMode = "Not Applicable"
			Else
				For Each objWMI In colWMI
'					WScript.Echo objWMI.State
'					WScript.Echo objWMI.Started
'					WScript.Echo objWMI.StartMode
					strState = UCase(objWMI.State)
					blnStarted = objWMI.Started
					strStartMode = objWMI.StartMode
				Next
			End If
		End If
	End If
	If ((strState = "") Or (strStartMode = "")) Then
		'
		' WMI isn't working - use SC command to get service status
		'
		strCommand = "SC \\" & strPassedParameter & " Query " & strServiceName
		intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, g_rsGeneric, g_objLogAndTraceExecCmdGeneric)
		If (intRetVal = 0) Then
			If (g_rsGeneric.RecordCount > 0) Then
				g_rsGeneric.MoveFirst
				While Not g_rsGeneric.EOF
					strLine = g_rsGeneric("SavedData")
					If (InStr(1, strLine, "STATE", vbTextCompare) > 0) Then
						If (InStr(1, strLine, "Stopped", vbTextCompare) > 0) Then
							strState = "Stopped"
						ElseIf (InStr(1, strLine, "Start_Pending", vbTextCompare) > 0) Then
							strState = "Start Pending"
						ElseIf (InStr(1, strLine, "Stop_Pending", vbTextCompare) > 0) Then
							strState = "Stop Pending"
						ElseIf (InStr(1, strLine, "Running", vbTextCompare) > 0) Then
							strState = "Running"
							blnStarted = True
						ElseIf (InStr(1, strLine, "Continue_Pending", vbTextCompare) > 0) Then
							strState = "Continue Pending"
						ElseIf (InStr(1, strLine, "Pause_Pending", vbTextCompare) > 0) Then
							strState = "Pause Pending"
						ElseIf (InStr(1, strLine, "Paused", vbTextCompare) > 0) Then
							strState = "Paused"
						Else
							strState = "Unknown"
						End If
					ElseIf (InStr(1, strLine, "The specified service does not exist as an installed service.", vbTextCompare) > 0) Then
						strState = "Not Installed"
					End If
					g_rsGeneric.MoveNext
				Wend
			End If
		End If
		g_objFunctions.DeleteAllRecordsetRows(g_rsGeneric)
		strCommand = "SC \\" & strPassedParameter & " QC " & strServiceName
		intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, g_rsGeneric, g_objLogAndTraceExecCmdGeneric)
		If (intRetVal = 0) Then
			If (g_rsGeneric.RecordCount > 0) Then
				g_rsGeneric.MoveFirst
				While Not g_rsGeneric.EOF
					strLine = g_rsGeneric("SavedData")
					If (InStr(1, strLine, "START_TYPE", vbTextCompare) > 0) Then
						If (InStr(1, strLine, "Boot", vbTextCompare) > 0) Then
							strStartMode = "Boot"
						ElseIf (InStr(1, strLine, "System", vbTextCompare) > 0) Then
							strStartMode = "System"
						ElseIf (InStr(1, strLine, "AUTO_START", vbTextCompare) > 0) Then
							strStartMode = "Auto"
						ElseIf (InStr(1, strLine, "DEMAND_START", vbTextCompare) > 0) Then
							strStartMode = "Manual"
						ElseIf (InStr(1, strLine, "Disabled", vbTextCompare) > 0) Then
							strStartMode = "Disabled"
						Else
							strStartMode = "Unknown"
						End If
					ElseIf (InStr(1, strLine, "The specified service does not exist as an installed service.", vbTextCompare) > 0) Then
						strStartMode = "Not Applicable"
					End If
					g_rsGeneric.MoveNext
				Wend
			End If
		End If
		g_objFunctions.DeleteAllRecordsetRows(g_rsGeneric)
	End If

End Function

Function CorrectCertificate(ByRef objRemoteRegServer, ByVal blnWMIRegGoodToGo, ByVal blnRemoteRegGoodToGo, ByVal strConnectedWithThis, _
								ByVal blnIs64BitMachine, ByVal strOSVersion, ByVal strCertificateToCorrect, ByVal intCertNumber, _
								ByVal strRegKey, ByRef objRepairFile)
'*****************************************************************************************************************************************
'*  Purpose:				Gets the certificates from the 
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				GetClientConfiguration
'*  Calls:					EnumRegistryKeys, GetRegistryEntry, RegKeySubkeyExists, DeleteRegistryKey, ExecCmdGeneric
'*							DeleteAllRecordsetRows, GetCertificateInfo
'*  Requirements:			None
'*****************************************************************************************************************************************
	Dim rsKeys, strEncryptionCert, strEncryption, intCount, strSigningCert, strSigning, strRegistryHive, strRegistryKey
	Dim blnKeyOrValueExists, strSearchValue, blnIsWow6432Node, valRegValue, strKeyType, blnDeletionRequired, blnProcessed, blnISaySo
	Dim intRetVal, strCommand
	
	Set rsKeys = CreateObject("ADODB.Recordset")

	If (strRegKey <> "FFFFFFFF") Then
		If ((blnWMIRegGoodToGo) Or (blnRemoteRegGoodToGo)) Then
			strRegistryHive = "HKLM"
			'
			' Delete the Certificate
			'
			objRepairFile.WriteLine("Deleting SMS " & strCertificateToCorrect & " Certificate from registry")
			Call g_objRegistryProcessing.DeleteRegistryKey(objRemoteRegServer, strConnectedWithThis, blnWMIRegGoodToGo, _
																blnRemoteRegGoodToGo, strRegistryHive, strRegKey, _
																blnIs64BitMachine, blnIsWow6432Node, "", "")
			'
			' Make sure the registry entry was deleted
			'
			blnISaySo = True
			intCount = 0
			While blnISaySo And intCount < 10
				objRepairFile.WriteLine("Ensuring SMS " & strCertificateToCorrect & " Certificate was deleted from registry")
				intRetVal = g_objRegistryProcessing.RegKeySubkeyExists(objRemoteRegServer, strRegistryHive, strRegKey)
				If (intRetVal = 0) Then
					blnProcessed = True
					blnISaySo = False
				Else
					objRepairFile.WriteLine("Deleting SMS " & strCertificateToCorrect & " Certificate from registry again (it still existed)")
					Call g_objRegistryProcessing.DeleteRegistryKey(objRemoteRegServer, strConnectedWithThis, blnWMIRegGoodToGo, _
																		blnRemoteRegGoodToGo, strRegistryHive, strRegKey, _
																		blnIs64BitMachine, blnIsWow6432Node, "", "")
					intCount = intCount + 1
					'
					' Wait 2 seconds prior to checking again
					'
					WScript.Sleep 2000
				End If
			Wend
			If (blnISaySo) Then
				objRepairFile.WriteLine("SMS " & strCertificateToCorrect & " Certificate wasn't deleted from registry after " & intCount & " validations.")
			Else
				objRepairFile.WriteLine("SMS " & strCertificateToCorrect & " Certificate was successfully deleted from the registry.")
			End If
		End If
	End If
		
	If (intCertNumber <> -1) Then
		'
		' The following code may only work if we are running on the remote computer.  We have found instances
		' when Certutil cannot open the certificate store due to issues with the mapping from machine to machine.
		' Since this occurs for other commands it isn't a certutil issue but a network issue.
		'
		objRepairFile.WriteLine("Deleting SMS " & strCertificateToCorrect & " Certificate")
		strCommand = "certutil -delstore \\" & strConnectedWithThis & "\SMS " & intCertNumber
		objRepairFile.WriteLine(strCommand)
		intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, g_rsGeneric, g_objLogAndTraceExecCmdGeneric)
		If (g_rsGeneric.RecordCount > 0) Then
			g_rsGeneric.MoveFirst
			While Not g_rsGeneric.EOF
				objRepairFile.WriteLine(g_rsGeneric("SavedData"))
				g_rsGeneric.MoveNext
			Wend
		End If
		g_objFunctions.DeleteAllRecordsetRows(g_rsGeneric)
		'
		' Make sure the Certificate has been removed from the store.
		'
		blnISaySo = True
		intCount = 0
		While blnISaySo And intCount < 10
			'
			' See if the SMS Certificate exists
			'
			objRepairFile.WriteLine("Ensuring SMS " & strCertificateToCorrect & " Certificate was deleted")
			strCommand = "certutil -store \\" & strConnectedWithThis & "\SMS " & intCertNumber
			intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, g_rsGeneric, g_objLogAndTraceExecCmdGeneric)
			If (intRetVal = 0) Then
				If (g_rsGeneric.RecordCount > 2) Then
					objRepairFile.WriteLine("Deleting SMS " & strCertificateToCorrect & " Certificate again (it still existed)")
					strCommand = "certutil -delstore \\" & strConnectedWithThis & "\SMS " & intCertNumber
					objRepairFile.WriteLine(strCommand)
					intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, g_rsGeneric, g_objLogAndTraceExecCmdGeneric)
					g_objFunctions.DeleteAllRecordsetRows(g_rsGeneric)
				Else
					blnISaySo = False
				End If
			End If
			intCount = intCount + 1
			'
			' Wait 2 seconds prior to checking again
			'
			WScript.Sleep 2000
		Wend
		If (blnISaySo) Then
			objRepairFile.WriteLine("SMS " & strCertificateToCorrect & " Certificate wasn't deleted after " & intCount & " validations.")
		Else
			objRepairFile.WriteLine("SMS " & strCertificateToCorrect & " Certificate was successfully deleted.")
		End If
	End If
	'
	' Cleanup
	'
	Set rsKeys = Nothing

End Function

Function ParseCertutilCertificateSMS(ByRef rsGeneric, ByRef intCertNumber, ByVal strType)
'*****************************************************************************************************************************************
'*  Purpose:				Loads Certificate information into XML file
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				All
'*  Calls:					CreateSetAndLinkAttribute, AppendChild
'*  Requirements:			None
'*****************************************************************************************************************************************
	Dim strLine, strCertHash, strSerialNumber, strIssuer, dtNotBefore, dtNotAfter, strSubject

	'
	' Initialize variables
	'
	strCertHash = ""
	strSerialNumber = ""
	strIssuer = ""
	dtNotBefore = DEFAULT_DATE
	dtNotAfter = DEFAULT_DATE
	strSubject = ""

	If (Not rsGeneric.BOF) Then
		rsGeneric.MoveFirst
	End If
	While Not rsGeneric.EOF
		strLine = rsGeneric("SavedData")
' 		'
' 		' Get the Certificate Number from the header
' 		'
'  		If (InStr(strLine, "= Certificate ") > 0) Then
'  			intCertNumber = Trim(Replace(strLine, "=", ""))
'  			intCertNumber = Trim(Replace(intCertNumber, "Certificate", ""))
'  		End If
		'
		' Get the Certificate SHA1 Hash
		'
		If (InStr(1, strLine, "Cert Hash(sha1):", vbTextCompare) > 0) Then
			strCertHash = Replace(Replace(strLine, "Cert Hash(sha1):", ""), " ", "")
		End If
		'
		' Get the Certificate Serial Number
		'
		If (InStr(1, strLine, "Serial Number:", vbTextCompare) > 0) Then
			strSerialNumber = Trim(Replace(strLine, "Serial Number:", ""))
		End If
		'
		' Get the Issuer of the Certificate
		'				
		If (InStr(1, strLine, "Issuer:", vbTextCompare) > 0) Then
			rsGeneric.MoveNext
			rsGeneric.MoveNext
			strIssuer = "CN=SMS, " & rsGeneric("SavedData")
			If (InStr(1, strIssuer, "OID.1.2.840.113549.1.9.2", vbTextCompare) > 0) Then
				strIssuer = Replace(strIssuer, "OID.1.2.840.113549.1.9.2", "VMWareCert")
			End If
		End If
		'
		' Get the Issue Date of the Certificate
		'
		If (InStr(1, strLine, "NotBefore:", vbTextCompare) > 0) Then
			dtNotBefore = Trim(Replace(strLine, "NotBefore:", ""))
			dtNotBefore = g_objFunctions.MassageTimestamp(dtNotBefore)
		End If
		'
		' Get the Expiration Date of the Certificate
		'
		If (InStr(1, strLine, "NotAfter:", vbTextCompare) > 0) Then
			dtNotAfter = Trim(Replace(strLine, "NotAfter:", ""))
			dtNotAfter = g_objFunctions.MassageTimestamp(dtNotAfter)
		End If
		'
		' Get the Subject of the Certificate
		'
		If (InStr(1, strLine, "Subject:", vbTextCompare) > 0) Then
			rsGeneric.MoveNext
			strSubject = "CN=SMS, " & rsGeneric("SavedData")
			If (InStr(1, strSubject, "OID.1.2.840.113549.1.9.2", vbTextCompare) > 0) Then
				strSubject = Replace(strSubject, "OID.1.2.840.113549.1.9.2", "VMWareCert")
			End If
		End If
		rsGeneric.MoveNext
	Wend
	If (strCertHash <> "") Then
		g_rsCertificates.AddNew
		Call g_objFunctions.LoadRS(g_rsCertificates, "SHA1Hash", strCertHash, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "SerialNumber", strSerialNumber, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "Issuer", strIssuer, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "NotBefore", dtNotBefore, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "NotAfter", dtNotAfter, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "Subject", strSubject, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "CertNumber", intCertNumber, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "RegistryKey", "FFFFFFFF", "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "Type", strType, "", "")
		g_rsCertificates.Update
	End If

End Function

Function ParseRegistryCertificate(ByVal valRegValue, ByVal strRegistryKey, ByVal strType)
'*****************************************************************************************************************************************
'*  Purpose:				Processes the Certificate Registry Blob and loads g_rsCertificates
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				GetCertificateSettings
'*  Calls:					HexToDec, WMIDateStringToDate
'*  Requirements:			None
'*****************************************************************************************************************************************
	Dim strPart1, strPart2, strWork, intLen, strCertHash, strSerialNumber, blnISaySo, intNextDC, intNextCN, strElement, intCount
	Dim strIssuer, dtTemp, dtNotBefore, dtNotAfter, intNextOU, strSubject

	'
	' Initialize variables
	'
	strCertHash = ""
	strSerialNumber = ""
	strElement = ""
	strIssuer = ""
	dtNotBefore = DEFAULT_DATE
	dtNotAfter = DEFAULT_DATE
	strSubject = ""
	'
	' The order of items in the certificate structure is as follows:
	'	Cert Hash
	'	Serial Number
	'	Issuer
	'	NotBefore
	'	NotAfter
	'	Subject
	'
	' Divide the string into 2 parts ->
	'	1. Cert Hash, Serial Number, and Issuer
	'	2. NotBefore, NotAfter, Subject
	'
	If ((InStr(1, valRegValue, "301E17", vbTextCompare) > 0) Or (InStr(1, valRegValue, "302017", vbTextCompare) > 0)) Then
		If (InStr(1, valRegValue, "302017", vbTextCompare) > 0) Then
			strPart1 = Split(valRegValue, "302017", 2, vbTextCompare)(0)
			strPart2 = Split(valRegValue, "302017", 2, vbTextCompare)(1)
		Else
			strPart1 = Split(valRegValue, "301E17", 2, vbTextCompare)(0)
			strPart2 = Split(valRegValue, "301E17", 2, vbTextCompare)(1)
		End If
		'
		' Process Certificate Hash
		'	Example:
		'		0300000001000000 (Start of Certificate Hash)
		'		14 (Certificate Hash length in Hex)
		'		000000
		'		67CD10DC8AB0E6431D7C2B010F48BE508B07BB69 (Cert Hash)
		'
		strWork = Split(strPart1, "0300000001000000", 2, vbTextCompare)(1)
		intLen = g_objFunctions.HexToDec(Mid(strWork, 1, 2))
		strCertHash = Mid(strWork, 9, intLen * 2)
' 		WScript.Echo "Cert Hash(SHA1): " & strCertHash
' 		WScript.Echo ""
		'
		' Process Serial Number
		'	Example:
		'		A00302010202 (Start of Serial Number)
		'		0A (Serial Number length in Hex)
		'		27090B6B0000000EE12E (Serial Number)
		'
		strWork = Split(strWork, "A00302010202", 2, vbTextCompare)(1)
		intLen = g_objFunctions.HexToDec(Mid(strWork, 1, 2))
		strSerialNumber = Mid(strWork, 3, intLen * 2)
' 		WScript.Echo "Serial Number: " & strSerialNumber
' 		WScript.Echo ""
		'
		' Process the Issuer
		'	Example:
		'		300D06092A864886F70D010105050030 (Start of Issuer)
		'		7E31133011 (different on every machine but the same number of characters)
		'		060A0992268993F22C64011916 (Indicates a "DC=")
		'		03 (Issuer part length in Hex)
		'		4D494C (MIL)
		'		31143012
		'		060A0992268993F22C64011916 (Indicates a "DC=")
		'		04 (Issuer part length in Hex)
		'		55534146 (USAF)
		'		31183016
		'		060A0992268993F22C64011916 (Indicates a "DC=")
		'		08 (Issuer part length in Hex)
		'		41464E4F41505053 (AFNOAPPS)
		'		31163014
		'		060A0992268993F22C64011916 (Indicates a "DC=")
		'		06 (Issuer part length in Hex)
		'		415245413532 (AREA52)
		'		311F301D
		'		060355040313 (Indicates a "CN=")
		'		16 (Issuer part length in Hex)
		'		555341462041464E4F41505053204D65642043412D33 (USAF AFNOAPPS MED CA-3)
		'		301E17 (Start of valid date section)
		'
		blnISaySo = True
		While blnISaySo
			intNextDC = InStr(1, strWork, "060a0992268993f22c64011916", vbTextCompare)
			intNextCN = InStr(1, strWork, "060355040313", vbTextCompare)
' 			WScript.Echo intNextDC
' 			WScript.Echo intNextCN
			If ((intNextDC = 0) And (intNextCN = 0)) Then
				blnISaySo = False
			Else
				If ((intNextDC <> 0) And (intNextDC < intNextCN)) Then
					'
					' The value for DC= appears first in the text string.  Parse it.
					'
					strWork = Right(strWork, Len(strWork) - (intNextDC + 26) + 1)
'					WScript.Echo "strWork: " & strWork
					intLen = g_objFunctions.HexToDec(Mid(strWork, 1, 2))
' 					WScript.Echo "intLen: " & intLen
					strElement = ""
					For intCount = 0 To intLen - 1
						strElement = strElement & Chr(g_objFunctions.HexToDec(Mid(strWork, (intCount * 2) + 3, 2)))
					Next
					If (strIssuer = "") Then
						strIssuer = "DC=" & strElement
					Else
						strIssuer = "DC=" & strElement & "," & strIssuer
					End If
' 					WScript.Echo "intLen: " & intLen
' 					WScript.Echo "Len(strWork): " & Len(strWork)
' 					strWork = Right(strWork, Len(strWork) - (intLen * 2) - 2)
' 					WScript.Echo strWork
				Else
					'
					' The value for CN= appears first in the text string.  Parse it.
					'
					strWork = Right(strWork, Len(strWork) - (intNextCN + 12) + 1)
'					WScript.Echo strWork
					intLen = g_objFunctions.HexToDec(Mid(strWork, 1, 2))
'					WScript.Echo "intLen: " & intLen
					strElement = ""
					For intCount = 0 To intLen - 1
						strElement = strElement & Chr(g_objFunctions.HexToDec(Mid(strWork, (intCount * 2) + 3, 2)))
					Next
					If (strIssuer = "") Then
						strIssuer = "CN=" & strElement
					Else
						strIssuer = "CN=" & strElement & "," & strIssuer
					End If
' 					strWork = Right(strWork, Len(strWork) - intLen - 14)
' 					WScript.Echo strWork
				End If
			End If
		Wend
' 		WScript.Echo "Issuer: " & strIssuer
' 		WScript.Echo ""
		'
		' Process NotBefore date
		'	Example:
		'		301E17 (Start of valid date section)
		'		0D (NotBefore length in Hex)
		'		3133303631353132343430395A (Not Before)
		'		17
		'		0D (NotAfter length in Hex)
		'		3134303631353132343430395A (Not After)
		'
		intLen = g_objFunctions.HexToDec(Mid(strPart2, 1, 2))
		dtTemp = ""
		If (intLen < 15) Then
			'
			' The year is in the same Century - append the first 2 digits of the year
			'
			dtTemp = Mid(Year(Now()), 1, 2)
		End If
		'
		' Skip the "Z" (Zulu) indicator at the end
		'
		For intCount = 0 To intLen
			dtTemp = dtTemp & Chr(g_objFunctions.HexToDec(Mid(strPart2, (intCount * 2) + 3, 2)))
		Next
		dtNotBefore = g_objFunctions.WMIDateStringToDate(dtTemp)
' 		WScript.Echo "Not Before: " & dtNotBefore
' 		WScript.Echo ""
		'
		' Move to the NotAfter date field
		'
		strWork = Right(strPart2, Len(strPart2) - (intLen * 2 + 4))
		'
		' Process NotAfter date
		'
		intLen = g_objFunctions.HexToDec(Mid(strWork, 1, 2))
		dtTemp = ""
		If (intLen < 15) Then
			'
			' The year is in the same Century - append the first 2 digits of the year
			'
			dtTemp = Mid(Year(Now()), 1, 2)
		End If
		'
		' Skip the "Z" (Zulu) indicator at the end
		'
		For intCount = 0 To intLen
			dtTemp = dtTemp & Chr(g_objFunctions.HexToDec(Mid(strWork, (intCount * 2) + 3, 2)))
		Next
		dtNotAfter = g_objFunctions.WMIDateStringToDate(dtTemp)
' 		WScript.Echo "Not After: " & dtNotAfter
' 		WScript.Echo ""
		'
		' Process Subject
		'	Example:
		'		3081CF31133011
		'		060A0992268993F22C64011916 (Indicates a "DC=")
		'		03 (Subject part length in Hex)
		'		4D494C (MIL)
		'		31143012
		'		060A0992268993F22C64011916 (Indicates a "DC=")
		'		04 (Subject part length in Hex)
		'		55534146 (USAF)
		'		31183016
		'		060A0992268993F22C64011916 (Indicates a "DC=")
		'		08 (Subject part length in Hex)
		'		41464E4F41505053 (AFNOAPPS)
		'		31163014
		'		060A0992268993F22C64011916 (Indicates a "DC=")
		'		06 (Subject part length in Hex)
		'		415245413532 (AREA52)
		'		310E300C
		'		060355040B13 (Indicates a "OU=")
		'		05 (Subject part length in Hex)
		'		4261736573 (Bases)
		'		31143012
		'		060355040B13 (Indicates a "OU=")
		'		0B (Subject part length in Hex)
		'		4146434F4E55535745535431123010 (AFCONUSWEST)
		'		060355040B13 (Indicates a "OU=")
		'		09 (Subject part length in Hex)
		'		53636F747420414642 (Scott AFB)
		'		311C301A
		'		060355040B13 (Indicates a "OU=")
		'		13 (Subject part length in Hex)
		'		53636F74742041464220436F6D707574657273 (Scott AFB Computers)
		'		31183016
		'		060355040313 (Indicates a "CN=")
		'		0F (Subject part length in Hex)
		'		3532564459444C332D534541303233 (52VDYDL3-SEA023)
		'
		'	Example 2:
		'		30233121301f
		'		060355040313 (Indicates a "CN=")
		'		18 (Subject part length in Hex)
		'		5048494c444330322e7068696c2e726f6f742e6c6f63616c (PHILDC02.phil.root.local)
		'
		blnISaySo = True
		While blnISaySo
' 			WScript.Echo "strWork: " & strWork
' 			WScript.Echo ""
			intNextDC = InStr(1, strWork, "060a0992268993f22c64011916", vbTextCompare)
			intNextOU = InStr(1, strWork, "060355040B13", vbTextCompare)
			intNextCN = InStr(1, strWork, "060355040313", vbTextCompare)
' 			WScript.Echo intNextDC
' 			WScript.Echo intNextOU
' 			WScript.Echo intNextCN
' 			WScript.Echo ""
			If ((intNextDC = 0) And (intNextOU = 0) And (intNextCN = 0)) Then
				blnISaySo = False
			Else
				If ((intNextDC <> 0) And (intNextDC < intNextOU) And (intNextDC < intNextCN)) Then
' 					WScript.Echo "Processing DC"
					'
					' The value for DC= appears first in the text string.  Parse it.
					'
					strWork = Right(strWork, Len(strWork) - (intNextDC + 26) + 1)
'					WScript.Echo "strWork: " & strWork
					intLen = g_objFunctions.HexToDec(Mid(strWork, 1, 2))
' 					WScript.Echo "intLen: " & intLen
					strElement = ""
					For intCount = 0 To intLen - 1
						strElement = strElement & Chr(g_objFunctions.HexToDec(Mid(strWork, (intCount * 2) + 3, 2)))
					Next
					If (strSubject = "") Then
						strSubject = "DC=" & strElement
					Else
						strSubject = "DC=" & strElement & "," & strSubject
					End If
' 					WScript.Echo "strSubject: " & strSubject
' 					WScript.Echo "Len(strWork): " & Len(strWork)
' 					strWork = Right(strWork, Len(strWork) - (intLen * 2) - 2)
' 					WScript.Echo strWork
				ElseIf ((intNextOU <> 0) And (intNextOU < intNextCN)) Then
' 					WScript.Echo "Processing OU"
					'
					' The value for OU= appears first in the text string.  Parse it.
					'
					strWork = Right(strWork, Len(strWork) - (intNextOU + 12) + 1)
'					WScript.Echo "strWork: " & strWork
					intLen = g_objFunctions.HexToDec(Mid(strWork, 1, 2))
' 					WScript.Echo "intLen: " & intLen
					strElement = ""
					For intCount = 0 To intLen - 1
						strElement = strElement & Chr(g_objFunctions.HexToDec(Mid(strWork, (intCount * 2) + 3, 2)))
					Next
					If (strSubject = "") Then
						strSubject = "OU=" & strElement
					Else
						strSubject = "OU=" & strElement & "," & strSubject
					End If
' 					WScript.Echo "intLen: " & intLen
' 					WScript.Echo "Len(strWork): " & Len(strWork)
' 					strWork = Right(strWork, Len(strWork) - (intLen * 2) - 2)
' 					WScript.Echo strWork
				Else
' 					WScript.Echo "Processing CN"
					'
					' The value for CN= appears first in the text string.  Parse it.
					'
					strWork = Right(strWork, Len(strWork) - (intNextCN + 12) + 1)
'					WScript.Echo strWork
					intLen = g_objFunctions.HexToDec(Mid(strWork, 1, 2))
'					WScript.Echo "intLen: " & intLen
					strElement = ""
					For intCount = 0 To intLen - 1
						strElement = strElement & Chr(g_objFunctions.HexToDec(Mid(strWork, (intCount * 2) + 3, 2)))
					Next
					If (strSubject = "") Then
						strSubject = "CN=" & strElement
					Else
						strSubject = "CN=" & strElement & "," & strSubject
					End If
' 					strWork = Right(strWork, Len(strWork) - intLen - 14)
' 					WScript.Echo strWork
				End If
' 				WScript.Echo "strSubject: " & strSubject
			End If
		Wend
' 		WScript.Echo "Subject: " & strSubject
' 		WScript.Echo ""
	End If
	If (strCertHash <> "") Then
		g_rsCertificates.AddNew
		Call g_objFunctions.LoadRS(g_rsCertificates, "SHA1Hash", strCertHash, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "SerialNumber", strSerialNumber, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "Issuer", strIssuer, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "NotBefore", dtNotBefore, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "NotAfter", dtNotAfter, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "Subject", strSubject, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "CertNumber", -1, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "RegistryKey", strRegistryKey, "", "")
		Call g_objFunctions.LoadRS(g_rsCertificates, "Type", strType, "", "")
		g_rsCertificates.Update
	End If

End Function

Function GetCertificateSettingsSMS(ByRef objRemoteRegServer, ByVal blnWMIRegGoodToGo, ByVal blnRemoteRegGoodToGo, _
									ByVal strConnectedWithThis, ByVal blnIs64BitMachine, ByVal strOSVersion)
'*****************************************************************************************************************************************
'*  Purpose:				Gets the specified certificate from the registry or certificate store
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				GetClientConfiguration
'*  Calls:					EnumRegistryKeys, GetRegistryEntry, ParseRegistryCertificate, ExecCmdGeneric, ParseCertutilCertificateSMS
'*  Requirements:			None
'*****************************************************************************************************************************************
	Dim rsKeys, strEncryptionCert, strEncryption, intCount, strSigningCert, strSigning, blnProcessedEncryption, blnProcessedSigning
	Dim strRegistryHive, strRegistryKey, blnKeyOrValueExists, strSearchValue, blnIsWow6432Node, valRegValue, strKeyType, rsGeneric
	Dim strCommand, intRetVal, blnEncryption, blnSigning, strLine
	
	Set rsKeys = CreateObject("ADODB.Recordset")
	strEncryptionCert = ""
	strEncryption = "Encryption"
	For intCount = 0 To Len(strEncryption) - 1
		strEncryptionCert = strEncryptionCert & Hex(Asc(Mid(strEncryption, intCount + 1, 1))) & "00"
	Next

	strSigningCert = ""
	strSigning = "Signing"
	For intCount = 0 To Len(strSigning) - 1
		strSigningCert = strSigningCert & Hex(Asc(Mid(strSigning, intCount + 1, 1))) & "00"
	Next
	'
	' Initialize
	'
	blnProcessedEncryption = False
	blnProcessedSigning = False
	
	If ((blnWMIRegGoodToGo) Or (blnRemoteRegGoodToGo)) Then
		strRegistryHive = "HKLM"
		strRegistryKey = "SOFTWARE\Microsoft\SystemCertificates\SMS\Certificates\"
		Call g_objRegistryProcessing.EnumRegistryKeys(objRemoteRegServer, strConnectedWithThis, blnWMIRegGoodToGo, blnRemoteRegGoodToGo, _
															strRegistryHive, strRegistryKey, blnIs64BitMachine, strOSVersion, _
															blnKeyOrValueExists, rsKeys, g_objLogAndTrace, g_objLogAndTraceErrors, _
															g_objLogAndTraceLoadRS)
		If ((blnKeyOrValueExists) And (rsKeys.RecordCount > 0)) Then
			If (Not rsKeys.BOF) Then
				rsKeys.MoveFirst
			End If
			While Not rsKeys.EOF
				strRegistryKey = rsKeys("Key") & "\" & rsKeys("Subkey") & "\"
				strSearchValue = "Blob"
				'
				' The SMS certificates are loaded in both locations on a 64-bit machine
				'
				blnIsWow6432Node = False
				valRegValue = ""
				Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strConnectedWithThis, blnWMIRegGoodToGo, _
																blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchValue, _
																blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
																strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
				If (blnKeyOrValueExists) Then
					If (InStr(1, valRegValue, strEncryptionCert, vbTextCompare) > 0) Then
						Call ParseRegistryCertificate(valRegValue, rsKeys("Key") & "\" & rsKeys("Subkey"), "Encryption")
						blnProcessedEncryption = True
					End If
					If (InStr(1, valRegValue, strSigningCert, vbTextCompare) > 0) Then
						Call ParseRegistryCertificate(valRegValue, rsKeys("Key") & "\" & rsKeys("Subkey"), "Signing")
						blnProcessedSigning = True
					End If
				End If
				rsKeys.MoveNext
			Wend
		End If
	End If
	'
	' If we are here then Registry processing wasn't available or it was and it failed - attempt to use Certutil
	'
	Set rsGeneric = CreateObject("ADODB.Recordset")
	rsGeneric.Fields.Append "SavedData", adVarChar, 255
	rsGeneric.Open

	If ((blnProcessedEncryption = False) Or (blnProcessedSigning = False)) Then
		For intCount = 0 To 1
			'
			' We don't know which certificate is which until we get the data back so call it twice then process as needed.
			'
			blnEncryption = False
			blnSigning = False
			strCommand = "certutil -store -v \\" & strConnectedWithThis & "\SMS " & intCount
			intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, rsGeneric, g_objLogAndTraceExecCmdGeneric)
			If (intRetVal = 0) Then
				If (rsGeneric.RecordCount > 2) Then
					rsGeneric.MoveFirst
					While Not rsGeneric.EOF
						strLine = rsGeneric("SavedData")
						'
						' "KeySpec = 1" ' -- AT_KEYEXCHANGE
						' "KeySpec = 2" ' -- AT_SIGNATURE
						'
						If (InStr(1, strLine, "KeySpec = 1", vbTextCompare) > 0) Then
							blnEncryption = True
						End If
						If (InStr(1, strLine, "KeySpec = 2", vbTextCompare) > 0) Then
							blnSigning = True
						End If
						rsGeneric.MoveNext
					Wend
					rsGeneric.MoveFirst
					If ((blnProcessedEncryption = False) And (blnEncryption)) Then
						Call ParseCertutilCertificateSMS(rsGeneric, intCount, "Encryption")
					End If
					If ((blnProcessedSigning = False) And (blnSigning)) Then
						Call ParseCertutilCertificateSMS(rsGeneric, intCount, "Signing")
					End If
				End If
			End If
			g_objFunctions.DeleteAllRecordsetRows(rsGeneric)
		Next
	End If
	'
	' Cleanup
	'
	Set rsKeys = Nothing
	Set rsGeneric = Nothing

End Function

Function BuildClientHealthXML(ByVal strPassedParameter, ByVal dtStartTime)
'*****************************************************************************************************************************************
'*  Purpose:				Build the ClientHealth XML file.
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				Mainline
'*  Calls:					CreateNewTableElement, CreateAppendNewElement, CreateSetAndLinkAttribute, CreateNewElement, AppendChild
'*							GetGMTTimestamp, FormatAndSaveXMLFile
'*  Requirements:			None
'*****************************************************************************************************************************************
	Dim xmlDoc, strActivityGUID, strConnectionStatusGUID, dtEndTime, xmlElementConnectionStatus, xmlElementClientHealth, xmlElementRow
	Dim xmlElementTables, strOutputFile, xmlElementProcessingInfo

	'
	' Create XML variables
	'
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	'
	' Setup the default ActivityGUID
	'
	strActivityGUID = "FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF"
	strConnectionStatusGUID = g_objFunctions.CreateGloballyUniqueID()
	dtEndTime = g_objFunctions.GetGMTTimestamp()
	'
	' Create the Table Elements
	'
	Call g_objXMLProcessing.CreateNewTableElement(xmlDoc, xmlElementConnectionStatus, "ConnectionStatus", "SEOSAMDB", "dbo")
	Call g_objXMLProcessing.CreateNewTableElement(xmlDoc, xmlElementClientHealth, "ClientHealth", "SEOSAMDB", "dbo")
	'
	' Create the ConnectionStatus element and Row child element
	'
	If (g_rsConnectionStatus.RecordCount > 0) Then
		g_rsConnectionStatus.MoveFirst
		Call g_objXMLProcessing.CreateAppendNewElement(xmlDoc, xmlElementConnectionStatus, xmlElementRow, "Row")
		'
		' Load the collected data
		'
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstConnectionStatusGUID", strConnectionStatusGUID)
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstActivityGUID", strActivityGUID)
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstPassedParameter", g_rsConnectionStatus("PassedParameter"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstHostAlive", g_rsConnectionStatus("HostAlive"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstWMIDCOMConnectionSuccessful", g_rsConnectionStatus("WMIDCOMConnectionSuccessful"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstWMIConnectedWithThis", g_rsConnectionStatus("WMIConnectedWithThis"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstWMIConnectErrorOccurred", g_rsConnectionStatus("WMIConnectErrorOccurred"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstWMIConnectError", g_rsConnectionStatus("WMIConnectError"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstWMIProcessingGoodToGo", g_rsConnectionStatus("WMIProcessingGoodToGo"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstRegistryProcessingMethodUsed", g_rsConnectionStatus("RegistryProcessingMethodUsed"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstWMIRegistryDCOMConnectionSuccessful", g_rsConnectionStatus("WMIRegistryDCOMConnectionSuccessful"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstWMIRegistryConnectedWithThis", g_rsConnectionStatus("WMIRegistryConnectedWithThis"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstWMIRegistryConnectErrorOccurred", g_rsConnectionStatus("WMIRegistryConnectErrorOccurred"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstWMIRegistryConnectError", g_rsConnectionStatus("WMIRegistryConnectError"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstWMIRegistryProcessingGoodToGo", g_rsConnectionStatus("WMIRegistryProcessingGoodToGo"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstRemoteRegistryConnectedWithThis", g_rsConnectionStatus("RemoteRegistryConnectedWithThis"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstRemoteRegistryProcessingGoodToGo", g_rsConnectionStatus("RemoteRegistryProcessingGoodToGo"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstLocalAccountsConnectedWithThis", g_rsConnectionStatus("LocalAccountsConnectedWithThis"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstLocalAccountProcessingGoodToGo", g_rsConnectionStatus("LocalAccountProcessingGoodToGo"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstProcessingLocal", g_rsConnectionStatus("ProcessingLocal"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstIPAddressAvailabilityMethod", g_rsConnectionStatus("IPAddressAvailabilityMethod"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstPassedDNSHostName", g_rsConnectionStatus("PassedDNSHostName"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstPassedHostName", g_rsConnectionStatus("PassedHostName"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstPassedIPAddress", g_rsConnectionStatus("PassedIPAddress"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstResolvedDNSHostName", g_rsConnectionStatus("ResolvedDNSHostName"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstResolvedHostName", g_rsConnectionStatus("ResolvedHostName"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstResolvedNetBIOSName", g_rsConnectionStatus("ResolvedNetBIOSName"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstResolvedIPAddress", g_rsConnectionStatus("ResolvedIPAddress"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstDNSHostNameResolved", g_rsConnectionStatus("DNSHostNameResolved"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstHostNameResolved", g_rsConnectionStatus("HostNameResolved"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstIPAddressResolved", g_rsConnectionStatus("IPAddressResolved"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstNetBIOSNameResolved", g_rsConnectionStatus("NetBIOSNameResolved"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstAbortedProcessingText", g_rsConnectionStatus("AbortedProcessingText"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstComputerFQDN", g_rsConnectionStatus("ComputerFQDN"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstProcessingStartTime", dtStartTime)
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstProcessingEndTime", dtEndTime)
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstIsPartialCollection", g_rsConnectionStatus("IsPartialCollection"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstIsCurrent", 1)
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "cstCreationTimestamp", g_dtCreationTimestamp)
		'
		' Write Trace information
		'
		Call LogIt(xmlElementRow, g_objLogAndTrace)
		Set xmlElementRow = Nothing
	End If
	'
	' Create the ClientHealth element and Row child element
	'
	If (g_rsClientConfiguration.RecordCount > 0) Then
		g_rsClientConfiguration.MoveFirst
		Call g_objXMLProcessing.CreateAppendNewElement(xmlDoc, xmlElementClientHealth, xmlElementRow, "Row")
		'
		' Load the collected data
		'
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhConnectionStatusGUID", strConnectionStatusGUID)
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhActivityGUID", strActivityGUID)
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhComputerFQDN", g_rsClientConfiguration("ComputerFQDN"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhComputerName", g_rsClientConfiguration("ComputerName"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhIPAddress", g_rsClientConfiguration("IPAddress"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhOSName", g_rsClientConfiguration("OSName"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhOSBuildNumber", g_rsClientConfiguration("OSBuildNumber"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhLastBootupTime", g_rsClientConfiguration("LastBootupTime"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhOSBuildType", g_rsClientConfiguration("OSBuildType"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhOSType", g_rsClientConfiguration("OSType"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhOSTypeText", g_rsClientConfiguration("OSTypeText"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhProductType", g_rsClientConfiguration("ProductType"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhProductTypeText", g_rsClientConfiguration("ProductTypeText"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhOSVersion", g_rsClientConfiguration("OSVersion"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhAddressWidth", g_rsClientConfiguration("AddressWidth"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhPercentFreeC", g_rsClientConfiguration("PercentFreeC"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhManuallyAssignedSiteCode", g_rsClientConfiguration("ManuallyAssignedSiteCode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhGPOAssignedSiteCode", g_rsClientConfiguration("GPOAssignedSiteCode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhEncryptionSubject", g_rsClientConfiguration("EncryptionSubject"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhEncryptionSubjectMatch", g_rsClientConfiguration("EncryptionSubjectMatch"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhEncryptionNotBefore", g_rsClientConfiguration("EncryptionNotBefore"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhEncryptionNotAfter", g_rsClientConfiguration("EncryptionNotAfter"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhSigningSubject", g_rsClientConfiguration("SigningSubject"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhSigningSubjectMatch", g_rsClientConfiguration("SigningSubjectMatch"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhSigningNotBefore", g_rsClientConfiguration("SigningNotBefore"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhSigningNotAfter", g_rsClientConfiguration("SigningNotAfter"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhWUAUState", g_rsClientConfiguration("WUAUState"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhWUAUStartMode", g_rsClientConfiguration("WUAUStartMode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhCCMEXECState", g_rsClientConfiguration("CCMEXECState"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhCCMEXECStartMode", g_rsClientConfiguration("CCMEXECStartMode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhBITSState", g_rsClientConfiguration("BITSState"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhBITSStartMode", g_rsClientConfiguration("BITSStartMode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhWinMgmtState", g_rsClientConfiguration("WinMgmtState"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhWinMgmtStartMode", g_rsClientConfiguration("WinMgmtStartMode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhGPClientState", g_rsClientConfiguration("GPClientState"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhGPClientStartMode", g_rsClientConfiguration("GPClientStartMode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhCCMSetupState", g_rsClientConfiguration("CCMSetupState"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhCCMSetupStartMode", g_rsClientConfiguration("CCMSetupStartMode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhLanmanServerState", g_rsClientConfiguration("LanmanServerState"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhLanmanServerStartMode", g_rsClientConfiguration("LanmanServerStartMode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhRemoteProcedureCallState", g_rsClientConfiguration("RPCSSState"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhRemoteProcedureCallStartMode", g_rsClientConfiguration("RPCSSStartMode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhSMSTaskSequenceAgentState", g_rsClientConfiguration("SMSTaskSequenceAgentState"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhSMSTaskSequenceAgentStartMode", g_rsClientConfiguration("SMSTaskSequenceAgentStartMode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhRemoteRegistryState", g_rsClientConfiguration("RemoteRegistryState"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhRemoteRegistryStartMode", g_rsClientConfiguration("RemoteRegistryStartMode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhSCCMVersion", g_rsClientConfiguration("SCCMVersion"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhWMIGoodToGo", g_rsClientConfiguration("WMIGoodToGo"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhWMIRegGoodToGo", g_rsClientConfiguration("WMIRegGoodToGo"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhRemoteRegGoodToGo", g_rsClientConfiguration("RemoteRegGoodToGo"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhWindowsUpdateServer", g_rsClientConfiguration("WindowsUpdateServer"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhNameServers", g_rsClientConfiguration("NameServers"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhLastLoggedOnUser", g_rsClientConfiguration("LastLoggedOnUser"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhPendingReboot", g_rsClientConfiguration("PendingReboot"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhSMSUID", g_rsClientConfiguration("SMSUID"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhPreviousSMSUID", g_rsClientConfiguration("PreviousSMSUID"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhLastChangedSMSUID", g_rsClientConfiguration("LastChangedSMSUID"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhClientSite", g_rsClientConfiguration("ClientSite"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhCurrentMP", g_rsClientConfiguration("CurrentMP"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhADSiteName", g_rsClientConfiguration("ADSiteName"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhSDCVersion", g_rsClientConfiguration("SDCVersion"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhWUAVersion", g_rsClientConfiguration("WUAVersion"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhEnableDCOM", g_rsClientConfiguration("EnableDCOM"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhSCCMHealthy", g_rsClientConfiguration("SCCMHealthy"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhRunningAdvertisements", g_rsClientConfiguration("RunningAdvertisements"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhMostRecentSMSFolder", g_rsClientConfiguration("MostRecentSMSFolder"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhMostRecentSMSFolderDate", g_rsClientConfiguration("MostRecentSMSFolderDate"))
		'
		' HBSS settings
		'
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhAVDatDate", g_rsClientConfiguration("AVDatDate"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhCatalogVersionDate", g_rsClientConfiguration("CatalogVersionDate"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhSiteListName", g_rsClientConfiguration("SiteListName"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhSiteListIP", g_rsClientConfiguration("SiteListIP"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhSiteListPort", g_rsClientConfiguration("SiteListPort"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhEPORegistryName", g_rsClientConfiguration("EPORegistryName"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhEPORegistryIP", g_rsClientConfiguration("EPORegistryIP"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhEPORegistryPort", g_rsClientConfiguration("EPORegistryPort"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhAgentGUID", g_rsClientConfiguration("AgentGUID"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhMcFrameworkState", g_rsClientConfiguration("McFrameworkState"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhMcFrameworkStartMode", g_rsClientConfiguration("McFrameworkStartMode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhMcShieldState", g_rsClientConfiguration("McShieldState"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhMcShieldStartMode", g_rsClientConfiguration("McShieldStartMode"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhLastASCTime", g_rsClientConfiguration("LastASCTime"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhPropsVersionDate", g_rsClientConfiguration("PropsVersionDate"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhAgentWakeUpPort", g_rsClientConfiguration("AgentWakeUpPort"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhMAVersion", g_rsClientConfiguration("MAVersion"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhFWEnabled", g_rsClientConfiguration("FWEnabled"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhVSEVersion", g_rsClientConfiguration("VSEVersion"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhHIPSVersion", g_rsClientConfiguration("HIPSVersion"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhDLPVersion", g_rsClientConfiguration("DLPVersion"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhHBSSHealthy", g_rsClientConfiguration("HBSSHealthy"))
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhIsCurrent", True)
		Call g_objXMLProcessing.CreateSetAndLinkAttribute(xmlDoc, xmlElementRow, "clhCreationTimestamp", g_dtCreationTimestamp)
	End If
	'
	' Add child element to the Tables Element
	'
	Call g_objXMLProcessing.CreateNewElement(xmlDoc, xmlElementTables, "Tables")
	'
	' Add child elements to the Tables Element
	'
	If (xmlElementConnectionStatus.HasChildNodes) Then
		Call g_objXMLProcessing.AppendChild(xmlElementConnectionStatus, xmlElementTables)
	End If
	If (xmlElementClientHealth.HasChildNodes) Then
		Call g_objXMLProcessing.AppendChild(xmlElementClientHealth, xmlElementTables)
	End If
	strOutputFile = g_strXMLOutputPath & strPassedParameter & ".xml"
	'
	' Create the ProcessInfo Element and add Attributes
	'	
	Call g_objXMLProcessing.CreateNewElement(g_xmlDoc, xmlElementProcessingInfo, "ProcessingInfo")
	Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, xmlElementProcessingInfo, "CreateActivityOnRollup", True)
	Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, xmlElementProcessingInfo, "ProcessingStartTime", dtStartTime)
	Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, xmlElementProcessingInfo, "ProcessingEndTime", dtEndTime)
	Call g_objXMLProcessing.FormatAndSaveXMLFile(strOutputFile, xmlElementTables, g_xmlElementPassedParams, xmlElementProcessingInfo)
	'
	' Cleanup
	'
	Set xmlDoc = Nothing
	Set xmlElementTables = Nothing
	Set xmlElementConnectionStatus = Nothing
	Set xmlElementClientHealth = Nothing
	Set xmlElementProcessingInfo = Nothing

End Function

Function GetClientConfiguration(ByVal strPassedParameter, ByVal intFlag, ByRef strWMIConnectedWithThis, ByRef objRemoteWMIServer, _
									ByRef objRemoteRegServer, ByRef blnIsClientAlive)
'*****************************************************************************************************************************************
'*  Purpose:				Gets the SCCM configuration for the specified machine.
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				Mainline
'*  Calls:					DetermineParameterType, GetClientNetworkInfo, ValidateProcessingOpportunities, GetOSBits, ExecWMI
'*							MassageTimestamp, GetServiceSettings, GetLastLoggedOnUser, VerifyAndLoad, GetRegistryEntry, EnumRegistryKeys
'*							EnumRegistryEntriesByKey, ValidateEntryExists, GetCertificateSettingsSMS
'*  Requirements:			None
'*****************************************************************************************************************************************
	Dim blnWMIGoodToGo, blnWMIRegGoodToGo, blnRemoteRegGoodToGo, blnLocalAccountProcessingGoodToGo, blnPassedDNSHostName, blnPassedHostName
	Dim blnPassedIPAddress, strValidationErrorsFile, objValidationErrorsFile, strResolvedDNSHostName, strResolvedHostName, strResolvedIPAddress
	Dim strResolvedNetBIOSName, strResolutionType, blnDNSHostNameResolved, blnHostNameResolved, blnNetBIOSNameResolved, blnIPAddressResolved
	Dim blnResolved, strWMIRegistryConnectedWithThis, strRemoteRegConnectedWithThis, strLocalAccountsConnectedWithThis, blnWMIConnectErrorOccurred
	Dim strWMIConnectError, blnWMIRegConnectErrorOccurred, strWMIRegistryConnectError, blnCOMConnectionSuccessful, rsKeys, rsEntries
	Dim strComputerFQDN, strComputerName, strIPAddress, strOSName, strOSBuildNumber, dtLastBootupTime, strOSBuildType, intOSType, strOSType
	Dim intProductType, strProductType, strOSVersion, strOSBits, intPercentFreeSpace, strManuallyAssignedSiteCode, strGPOAssignedSiteCode
	Dim strEncryptionSubject, dtEncryptionNotBefore, dtEncryptionNotAfter, blnEncryptionSubjectMatch, intEncryptionCertNumber, strEncryptionRegKey
	Dim strSigningSubject, dtSigningNotBefore, dtSigningNotAfter, blnSigningSubjectMatch, intSigningCertNumber, strSigningRegKey, strWUAUState
	Dim strWUAUStartMode, strSCCMState, strSCCMStartMode, strBITSState, strBITSStartMode, strWinMgmtState, strWinMgmtStartMode, strMcFrameworkState
	Dim strMcFrameworkStartMode, strMcShieldState, strMcShieldStartMode, strMcFirewallCoreState, strMcFirewallCoreStartMode, strMcHIPSState
	Dim strMcHIPSStartMode, strMcDLPState, strMcDLPStartMode, strGPClientState, strGPClientStartMode, strSCCMVersion, strWindowsUpdateServer
	Dim strNameServers, dtAVDatDate, dtLastASCTime, dtPropsVersionDate, intAgentWakeUpPort, arrServerList, intCount, arrSplit,strEPORegistryName
	Dim strEPORegistryIP, strEPORegistryPort, strAgentGUID, strMAVersion, blnFWEnabled, strVSEVersion, strHIPSVersion, strDLPVersion
	Dim dtCatalogVersionDate, strSiteListName, strSiteListIP, strSiteListPort, strLastLoggedOnUser, blnPendingReboot, strSMSUID, strPreviousSMSUID
	Dim dtLastChangedSMSUID, strMachineSID, strSMBIOSSerialNumber, strHardwareIdentifier, strHardwareIdentifier2, strLastVersion, strClientSite
	Dim strCurrentMP, strADSiteName, strSDCVersion, strWUAVersion, strCCMSetupState, strCCMSetupStartMode, strLanmanServerState
	Dim strLanmanServerStartMode, strRPCSSState, strRPCSSStartMode, strSMSTSMGRState, strSMSTSMGRStartMode, strRemoteRegistryState
	Dim strRemoteRegistryStartMode, intRunningAdvertisements, blnEnableDCOM, blnSCCMHealthy, blnHBSSHealthy, blnIs64BitMachine, intOSBits
	Dim strSQLQuery, intErrNumber, strErrDescription, colWMI, objWMI, fltFreeSpace, fltTotalSpace, fltUsedSpace, blnStarted, strFrameworkAgentFile
	Dim strFrameworkINIFile, objFrameworkINIFile, blnISaySo, strLine, strCatalogVersionDate, rsSiteList, strSiteListFile, strSiteListXMLFile
	Dim xmlDoc, colRowNodes, objRowNode, colAttributes, objAttribute, strAttributeName, valAttributeValue, strServerName, intOrder, strServerIP
	Dim intSecurePort, strSMSConfigFile, strSiteSMSConfigFile, objSiteSMSConfigFile, strRegistryHive, strRegistryKey, strSearchEntry, valRegValue
	Dim blnKeyOrValueExists, strKeyType, strWindowsDirectory, strSystemDirectory, strSCCMCacheFolder, dtMostRecent, strMostRecent, colFolders
	Dim objSubFolder, colSubFolders, dtCreationDate, strFolderName, strSystemDirectoryForConnect, colFiles, objFile, strNameSpaceForConnection
	Dim strNameSpaceToValidate, strNameSpace, blnNamespaceInstalled, intRetVal, objSMS, strError, strSubkey, strNewRegistryKey, blnIsWow6432Node
	Dim blnEntryExists, blnDHCPEnabled, strIPAddressFromRegistry, dtThirtyDaysAgo, dtSevenDaysAgo

	'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion
	'    CurrentVersion    REG_SZ    6.0
	'    CurrentBuildNumber    REG_SZ    6002
	'    CurrentBuild    REG_SZ    6002
	'    SoftwareType    REG_SZ    System
	'    CurrentType    REG_SZ    Multiprocessor Free
	'    InstallDate    REG_DWORD    0x4c0fc012
	'    RegisteredOrganization    REG_SZ    U.S. Air Force
	'    RegisteredOwner    REG_SZ    U.S. Air Force User
	'    SystemRoot    REG_SZ    C:\Windows
	'    ProductName    REG_SZ    Windows Vista (TM) Enterprise
	'    ProductId    REG_SZ    89579-236-0200203-71984
	'    DigitalProductId    REG_BINARY    A40000000300000038393537392D3233362D303230
	'    DigitalProductId4    REG_BINARY    F804000004000000380039003500370039002D003
	'    EditionID    REG_SZ    Enterprise
	'    BuildLab    REG_SZ    6002.vistasp2_gdr.100218-0019
	'    BuildLabEx    REG_SZ    6002.18209.x86fre.vistasp2_gdr.100218-0019
	'    BuildGUID    REG_SZ    a88c5de1-11b5-4a92-b4bd-b045f921b4f7
	'    CSDBuildNumber    REG_SZ    1621
	'    PathName    REG_SZ    C:\Windows
	'    CSDVersion    REG_SZ    Service Pack 2
	'
	' HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment
	'	windir
	'
	blnWMIGoodToGo = False
	blnWMIRegGoodToGo = False
	blnRemoteRegGoodToGo = False
	blnLocalAccountProcessingGoodToGo = False
	'
	' Validate parameters
	'
	Call g_objFunctions.DetermineParameterType(strPassedParameter, blnPassedDNSHostName, blnPassedHostName, blnPassedIPAddress)
	If ((blnPassedDNSHostName = False) And (blnPassedHostName = False) And (blnPassedIPAddress = False)) Then
		'
		' Whatever was passed couldn't be determined (should never happen)
		'
		strValidationErrorsFile = g_strErrorOutputPath & "ValidationErrors_" & strPassedParameter & "_" & g_objFunctions.BuildDateString(Now) & ".txt"
		Set objValidationErrorsFile = g_objFSO.OpenTextFile(strValidationErrorsFile, FOR_WRITE, CREATE_IF_NON_EXISTENT)
		objValidationErrorsFile.WriteLine("You passed crap for " & strPassedParameter & " so we are done...skipping")
		objValidationErrorsFile.Close
		Set objValidationErrorsFile = Nothing
		GetClientConfiguration = -1
		Exit Function
	End If
	'
	' Is client alive for processing?
	'
	Call g_objClientResolution.GetClientNetworkInfo(strPassedParameter, blnPassedDNSHostName, blnPassedHostName, blnPassedIPAddress, _
														strResolvedDNSHostName, strResolvedHostName, strResolvedIPAddress, _
														strResolvedNetBIOSName, strResolutionType, blnIsClientAlive, blnDNSHostNameResolved, _
														blnHostNameResolved, blnNetBIOSNameResolved, blnIPAddressResolved, blnResolved, _
														g_objLogAndTraceClientResolution, g_strLogAndTraceExecCmdGeneric)
	'
	' What processing can we do?
	'
	Call g_objPossibleProcessing.ValidateProcessingOpportunities(strPassedParameter, "", "", strResolvedDNSHostName, strResolvedHostName, _
																	strResolvedIPAddress, strResolvedNetBIOSName, blnDNSHostNameResolved, _
																	blnHostNameResolved, blnNetBIOSNameResolved, blnIPAddressResolved, False, _
																	strWMIConnectedWithThis, objRemoteWMIServer, blnWMIGoodToGo, _
																	strWMIRegistryConnectedWithThis, objRemoteRegServer, blnWMIRegGoodToGo, _
																	strRemoteRegConnectedWithThis, blnRemoteRegGoodToGo, _
																	strLocalAccountsConnectedWithThis, blnLocalAccountProcessingGoodToGo, _
																	blnWMIConnectErrorOccurred, strWMIConnectError, blnWMIRegConnectErrorOccurred, _
																	strWMIRegistryConnectError, blnCOMConnectionSuccessful, g_objLogAndTrace, _
																	g_objLogAndTraceErrors)
	If ((blnWMIGoodToGo) Or (blnWMIRegGoodToGo) Or (blnRemoteRegGoodToGo)) Then
		blnIsClientAlive = True
	End If
	'
	' Create the recordset entry for the data collected
	'
	g_rsConnectionStatus.AddNew
	g_rsConnectionStatus("PassedParameter") = strPassedParameter
	g_rsConnectionStatus("HostAlive") = blnIsClientAlive
	g_rsConnectionStatus("WMIDCOMConnectionSuccessful") = blnCOMConnectionSuccessful
	g_rsConnectionStatus("WMIConnectedWithThis") = strWMIConnectedWithThis
	g_rsConnectionStatus("WMIConnectErrorOccurred") = blnWMIConnectErrorOccurred
	g_rsConnectionStatus("WMIConnectError") = strWMIConnectError
	g_rsConnectionStatus("WMIProcessingGoodToGo") = blnWMIGoodToGo
	If (blnWMIRegGoodToGo) Then
		g_rsConnectionStatus("RegistryProcessingMethodUsed") = "WMIRegistryProcessing"
	ElseIf (blnRemoteRegGoodToGo) Then
		g_rsConnectionStatus("RegistryProcessingMethodUsed") = "RemoteRegistryProcessing"
	Else
		g_rsConnectionStatus("RegistryProcessingMethodUsed") = "None"
	End If
	g_rsConnectionStatus("WMIRegistryDCOMConnectionSuccessful") = blnWMIRegConnectErrorOccurred
	g_rsConnectionStatus("WMIRegistryConnectedWithThis") = LCase(strWMIRegistryConnectedWithThis)
	g_rsConnectionStatus("WMIRegistryConnectErrorOccurred") = blnWMIRegConnectErrorOccurred
	g_rsConnectionStatus("WMIRegistryConnectError") = strWMIRegistryConnectError
	g_rsConnectionStatus("WMIRegistryProcessingGoodToGo") = blnWMIRegGoodToGo
	g_rsConnectionStatus("RemoteRegistryConnectedWithThis") = LCase(strRemoteRegConnectedWithThis)
	g_rsConnectionStatus("RemoteRegistryProcessingGoodToGo") = blnRemoteRegGoodToGo
	g_rsConnectionStatus("LocalAccountsConnectedWithThis") = LCase(strLocalAccountsConnectedWithThis)
	g_rsConnectionStatus("LocalAccountProcessingGoodToGo") = blnLocalAccountProcessingGoodToGo
	g_rsConnectionStatus("ProcessingLocal") = g_blnProcessingLocal
	g_rsConnectionStatus("IPAddressAvailabilityMethod") = strResolutionType
	g_rsConnectionStatus("PassedDNSHostName") = blnPassedDNSHostName
	g_rsConnectionStatus("PassedHostName") = blnPassedHostName
	g_rsConnectionStatus("PassedIPAddress") = blnPassedIPAddress
	g_rsConnectionStatus("ResolvedDNSHostName") = LCase(strResolvedDNSHostName)
	g_rsConnectionStatus("ResolvedHostName") = UCase(strResolvedHostName)
	g_rsConnectionStatus("ResolvedNetBIOSName") = strResolvedNetBIOSName
	g_rsConnectionStatus("ResolvedIPAddress") = strResolvedIPAddress
	g_rsConnectionStatus("DNSHostNameResolved") = blnDNSHostNameResolved
	g_rsConnectionStatus("HostNameResolved") = blnHostNameResolved
	g_rsConnectionStatus("IPAddressResolved") = blnIPAddressResolved
	g_rsConnectionStatus("NetBIOSNameResolved") = blnNetBIOSNameResolved
	g_rsConnectionStatus("AbortedProcessingText") = ""
	g_rsConnectionStatus("ComputerFQDN") = LCase(strResolvedDNSHostName)
	g_rsConnectionStatus("IsPartialCollection") = 0
	g_rsConnectionStatus.Update
	If (blnIsClientAlive = False) Then
		'
		' Machine is truly not alive
		'
		strValidationErrorsFile = g_strErrorOutputPath & "ValidationErrors_" & strPassedParameter & "_" & g_objFunctions.BuildDateString(Now) & ".txt"
		Set objValidationErrorsFile = g_objFSO.OpenTextFile(strValidationErrorsFile, FOR_WRITE, CREATE_IF_NON_EXISTENT)
		If (blnIsClientAlive = False) Then
			objValidationErrorsFile.WriteLine("Machine " & strPassedParameter & " is not alive...skipping")
		End If
		objValidationErrorsFile.Close
		Set objValidationErrorsFile = Nothing
		GetClientConfiguration = -1
		Exit Function
	ElseIf ((blnWMIGoodToGo = False) Or ((blnWMIRegGoodToGo = False) And (blnRemoteRegGoodToGo = False))) Then
		'
		' Client is alive - WMI Is not available
		'
		strValidationErrorsFile = g_strErrorOutputPath & "ValidationErrors_" & strPassedParameter & "_" & g_objFunctions.BuildDateString(Now) & ".txt"
		Set objValidationErrorsFile = g_objFSO.OpenTextFile(strValidationErrorsFile, FOR_WRITE, CREATE_IF_NON_EXISTENT)
		objValidationErrorsFile.WriteLine("Machine " & strPassedParameter & " WMI not available...skipping")
		objValidationErrorsFile.Close
		Set objValidationErrorsFile = Nothing
		GetClientConfiguration = -1
		Exit Function
	End If
	Set rsKeys = CreateObject("ADODB.Recordset")
	Set rsEntries = CreateObject("ADODB.Recordset")
	'
	' Initialize variables
	'
	strComputerFQDN = strResolvedDNSHostName
	strComputerName = strResolvedHostName
	If (g_blnProcessingLocal) Then
		Call g_objClientResolution.GetActiveIPAddress(objRemoteWMIServer, intFlag, strIPAddress, g_objLogAndTrace)
	Else
		strIPAddress = strResolvedIPAddress
	End If
	strOSName = ""
	strOSBuildNumber = ""
	dtLastBootupTime = ""
	strOSBuildType = ""
	intOSType = ""
	strOSType = ""
	intProductType = ""
	strProductType = ""
	strOSVersion = ""
	strOSBits = ""
	intPercentFreeSpace = -1
	strManuallyAssignedSiteCode = ""
	strGPOAssignedSiteCode = ""
	strEncryptionSubject = ""
	dtEncryptionNotBefore = DEFAULT_DATE
	dtEncryptionNotAfter = DEFAULT_DATE
	blnEncryptionSubjectMatch = False
	intEncryptionCertNumber = -1
	strEncryptionRegKey = "FFFFFFFF"
	strSigningSubject = ""
	dtSigningNotBefore = DEFAULT_DATE
	dtSigningNotAfter = DEFAULT_DATE
	blnSigningSubjectMatch = False
	intSigningCertNumber = -1
	strSigningRegKey = "FFFFFFFF"
	strWUAUState = "Unknown"
	strWUAUStartMode = "Unknown"
	strSCCMState = "Unknown"
	strSCCMStartMode = "Unknown"
	strBITSState = "Unknown"
	strBITSStartMode = "Unknown"
	strWinMgmtState = "Unknown"
	strWinMgmtStartMode = "Unknown"
	strMcFrameworkState = "Unknown"
	strMcFrameworkStartMode = "Unknown"
	strMcShieldState = "Unknown"
	strMcShieldStartMode = "Unknown"
	strMcFirewallCoreState = "Unknown"
	strMcFirewallCoreStartMode = "Unknown"
	strMcHIPSState = "Unknown"
	strMcHIPSStartMode = "Unknown"
	strMcDLPState = "Unknown"
	strMcDLPStartMode = "Unknown"
	strGPClientState = "Unknown"
	strGPClientStartMode = "Unknown"
	strSCCMVersion = ""
	strWindowsUpdateServer = ""
	strNameServers = ""
	dtAVDatDate = DEFAULT_DATE
	dtLastASCTime = DEFAULT_DATE
	dtPropsVersionDate = DEFAULT_DATE
	intAgentWakeUpPort = -1
	strEPORegistryName = ""
	strEPORegistryIP = ""
	strEPORegistryPort = ""
	strAgentGUID = ""
	strMAVersion = ""
	blnFWEnabled = False
	strVSEVersion = ""
	strHIPSVersion = ""
	strDLPVersion = ""
	dtCatalogVersionDate = DEFAULT_DATE
	strSiteListName = ""
	strSiteListIP = ""
	strSiteListPort = ""
	strLastLoggedOnUser = ""
	blnPendingReboot = False
	strSMSUID = ""
	strPreviousSMSUID = ""
	dtLastChangedSMSUID = DEFAULT_DATE
	strMachineSID = ""
	strSMBIOSSerialNumber = ""
	strHardwareIdentifier = ""
	strHardwareIdentifier2 = ""
	strLastVersion = ""
	strClientSite = "Unavailable"
	strCurrentMP = "Unavailable"
	strADSiteName = "Unavailable"
	strSDCVersion = ""
	strWUAVersion = ""
	strCCMSetupState = "Unknown"
	strCCMSetupStartMode = "Unknown"
	strLanmanServerState = "Unknown"
	strLanmanServerStartMode = "Unknown"
	strRPCSSState = "Unknown"
	strRPCSSStartMode = "Unknown"
	strSMSTSMGRState = "Unknown"
	strSMSTSMGRStartMode = "Unknown"
	strRemoteRegistryState = "Unknown"
	strRemoteRegistryStartMode = "Unknown"
	intRunningAdvertisements = -1
	blnEnableDCOM = False
	blnSCCMHealthy = False
	blnHBSSHealthy = False
	'
	' If we are here then WMI is good
	'
	blnIs64BitMachine = False
	'
	' Start WMI processing
	'
	If (blnWMIGoodToGo) Then
		'
		' Determine if the remote computer is a 32-bit OS or 64-bit OS
		'
		Call g_objFunctions.GetOSBits(objRemoteWMIServer, objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIGoodToGo, _
											blnWMIRegGoodToGo, blnRemoteRegGoodToGo, blnIs64BitMachine)	
		If (blnIs64BitMachine) Then
			intOSBits = 64
			strOSBits = "64"
		Else
			intOSBits = 32
			strOSBits = "32"
		End If
		strSQLQuery = "SELECT Name,BuildNumber,BuildType,LastBootupTime,OSType,ProductType,Version FROM Win32_OperatingSystem"
		Call g_objFunctions.ExecWMI(objRemoteWMIServer, intErrNumber, strErrDescription, colWMI, strSQLQuery, intFlag, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
			For Each objWMI In colWMI
				strOSName = objWMI.Name
				If (InStr(strOSName, "|")) Then
					strOSName = Split(strOSName, "|", 2)(0)
				End If
				strOSBuildNumber = objWMI.BuildNumber
				strOSBuildType = objWMI.BuildType
				dtLastBootupTime = objWMI.LastBootupTime
				If ((IsNull(dtLastBootupTime)) Or (IsEmpty(dtLastBootupTime))) Then
					dtLastBootupTime = DEFAULT_DATE
				Else
					dtLastBootupTime = CDate(g_objFunctions.WMIDateStringToDate(dtLastBootupTime))
				End If
				dtLastBootupTime = g_objFunctions.MassageTimestamp(dtLastBootupTime)
				intOSType = objWMI.OSType
				Select Case intOSType
					Case 14
						strOSType = "MSDOS"
					Case 15
						strOSType = "WIN3x"
					Case 16
						strOSType = "WIN95"
					Case 17
						strOSType = "WIN98"
					Case 18
						strOSType = "WINNT"
					Case 19
						strOSType = "WINCE"
					Case Else
						strOSType = "Non-Microsoft"
				End Select
				intProductType = objWMI.ProductType
				Select Case intProductType
					Case 1
						strProductType = "Workstation"
					Case 2
						strProductType = "Domain Controller"
					Case 3
						strProductType = "Server"
					Case Else
						strProductType = "Not Defined"
				End Select
				strOSVersion = Left(objWMI.Version, InStrRev(objWMI.Version, ".") - 1)
			Next
		End If
		strSQLQuery = "SELECT FreeSpace, Size FROM Win32_LogicalDisk WHERE DeviceID = 'C:'"
		Call g_objFunctions.ExecWMI(objRemoteWMIServer, intErrNumber, strErrDescription, colWMI, strSQLQuery, intFlag, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
			For Each objWMI In colWMI
				fltFreeSpace = objWMI.FreeSpace
				fltTotalSpace = objWMI.Size
				fltUsedSpace = objWMI.Size - objWMI.FreeSpace
				intPercentFreeSpace = Round(((fltFreeSpace / fltTotalSpace) * 100), 2)
			Next
		End If
	End If
	'
	' See if Windows Update service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "WUAUSERV", strPassedParameter, _
								strWUAUState, blnStarted, strWUAUStartMode)
	'
	' See if SMS service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "CCMEXEC", strPassedParameter, _
								strSCCMState, blnStarted, strSCCMStartMode)
	'
	' See if BITS service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "BITS", strPassedParameter, _
								strBITSState, blnStarted, strBITSStartMode)
	'
	' See if WinMgmt service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "WINMGMT", strPassedParameter, _
								strWinMgmtState, blnStarted, strWinMgmtStartMode)
	'
	' See if Group Policy Client service is running
	'
	If (Mid(strOSVersion, 1, 1) >= "6") Then
		'
		' Vista and above
		'
		Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "GPSVC", strPassedParameter, _
									strGPClientState, blnStarted, strGPClientStartMode)
	Else
		strGPClientState = "N/A"
		strGPClientStartMode = "N/A"
	End If
	'
	' See if CCMSetup service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "CCMSETUP", strPassedParameter, _
								strCCMSetupState, blnStarted, strCCMSetupStartMode)
	'
	' See if LanmanServer service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "LANMANSERVER", strPassedParameter, _
								strLanmanServerState, blnStarted, strLanmanServerStartMode)
	'
	' See if RPCSS service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "RPCSS", strPassedParameter, _
								strRPCSSState, blnStarted, strRPCSSStartMode)
	'
	' See if SMS Task Sequence Agent (SMSTSMGR) service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "SMSTSMGR", strPassedParameter, _
								strSMSTSMGRState, blnStarted, strSMSTSMGRStartMode)
	'
	' See if RemoteRegistry service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "REMOTEREGISTRY", strPassedParameter, _
								strRemoteRegistryState, blnStarted, strRemoteRegistryStartMode)
	'
	' HBSS Services:
	'
	'
	' See if McAfee Framework service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "MCAFEEFRAMEWORK", strPassedParameter, _
								strMcFrameworkState, blnStarted, strMcFrameworkStartMode)
	'
	' See if McAfee McShield service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "MCSHIELD", strPassedParameter, _
								strMcShieldState, blnStarted, strMcShieldStartMode)
	'
	' See if McAfee Firewall Core service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "MFEFIRE", strPassedParameter, _
								strMcFirewallCoreState, blnStarted, strMcFirewallCoreStartMode)
	'
	' See if McAfee HIPS service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "ENTERCEPTAGENT", strPassedParameter, _
								strMcHIPSState, blnStarted, strMcHIPSStartMode)
	'
	' See if McAfee DLP Endpoint Agent service is running
	'
	Call GetServiceSettings(objRemoteWMIServer, intFlag, blnWMIGoodToGo, "MCAFEEDLPAGENTSERVICE", strPassedParameter, _
								strMcDLPState, blnStarted, strMcDLPStartMode)
	'
	' Get the Catalog Version from the McAfee Framework Agent ini file (CatalogVersion=20131021010131)
	'
	If (Mid(strOSVersion, 1, 1) >= "6") Then
		strFrameworkAgentFile = "C:\ProgramData\McAfee\Common Framework\agent.ini"
	Else
		strFrameworkAgentFile = "c:\Documents And Settings\All Users\Application Data\McAfee\Common Framework\agent.ini"
	End If
	strFrameworkINIFile = "\\" & strPassedParameter & "\" & Replace(strFrameworkAgentFile, ":", "$")
	If (g_objFSO.FileExists(strFrameworkINIFile)) Then
		Set objFrameworkINIFile = g_objFSO.OpenTextFile(strFrameworkINIFile)
		blnISaySo = True
		While Not objFrameworkINIFile.AtEndOfStream And blnISaySo
			strLine = UCase(Trim(objFrameworkINIFile.ReadLine))
			If (InStr(1, strLine, "CatalogVersion", vbTextCompare) > 0) Then
				strCatalogVersionDate = Split(strLine, "=", 2)(1)
				If (Mid(strCatalogVersionDate, 1, 1) = "0") Then
					dtCatalogVersionDate = DEFAULT_DATE
				Else
					dtCatalogVersionDate = g_objFunctions.WMIDateStringToDate(strCatalogVersionDate)
				End If
				dtCatalogVersionDate = g_objFunctions.MassageTimestamp(dtCatalogVersionDate)
				blnISaySo = False
			End If
		Wend
	End If
	'
	' Get the current/last logged on user
	'
	Call GetLastLoggedOnUser(objRemoteWMIServer, intFlag, strLastLoggedOnUser)
	'
	' Create and open the recordset for SiteList
	'
	Set rsSiteList = CreateObject("ADODB.Recordset")
	rsSiteList.Fields.Append "Order", adInteger
	rsSiteList.Fields.Append "ServerName", adVarChar, 50
	rsSiteList.Fields.Append "ServerIP", adVarChar, 30
	rsSiteList.Fields.Append "SecurePort", adInteger
	rsSiteList.Open
	'
	' Get the HBSS Server information
	'
	If (Mid(strOSVersion, 1, 1) >= "6") Then
		strSiteListFile = "C:\ProgramData\McAfee\Common Framework\sitelist.xml"
	Else
		strSiteListFile = "c:\Documents And Settings\All Users\Application Data\McAfee\Common Framework\sitelist.xml"
	End If
	strSiteListXMLFile = "\\" & strPassedParameter & "\" & Replace(strSiteListFile, ":", "$")
	If (g_objFSO.FileExists(strSiteListXMLFile)) Then
		On Error Resume Next
		Set xmlDoc = CreateObject("Microsoft.XMLDOM")
		intErrNumber = Err.Number
		strErrDescription = Err.Description
		On Error GoTo 0
		If (intErrNumber = 0) Then
			xmlDoc.async = False
			xmlDoc.load(strSiteListXMLFile)		' Load the entire document before continuing processing
			xmlDoc.validateOnParse = False
			xmlDoc.preserveWhiteSpace = True
			xmlDoc.resolveExternals = False
			Set colRowNodes = xmlDoc.documentElement.selectNodes("//SiteList/SpipeSite")
			For Each objRowNode In colRowNodes
				Set colAttributes = objRowNode.Attributes
				For Each objAttribute In colAttributes
					strAttributeName = objAttribute.Name
					valAttributeValue = objAttribute.Value
					Select Case UCase(strAttributeName)
						Case "SERVER"
							strServerName = valAttributeValue
						Case "ORDER"
							intOrder = valAttributeValue
						Case "SERVERIP"
							strServerIP = valAttributeValue
						Case "SECUREPORT"
							intSecurePort = valAttributeValue
						Case Else
					End Select
				Next
				rsSiteList.AddNew
				rsSiteList("Order") = intOrder
				rsSiteList("ServerName") = strServerName
				rsSiteList("ServerIP") = strServerIP
				rsSiteList("SecurePort") = intSecurePort
				rsSiteList.Update
			Next
		End If
	End If
	If (rsSiteList.RecordCount > 0) Then
		rsSiteList.MoveFirst
		While Not rsSiteList.EOF
			intOrder = rsSiteList("Order")
			strServerName = rsSiteList("ServerName")
			strServerIP = rsSiteList("ServerIP")
			intSecurePort = rsSiteList("SecurePort")
			If (strSiteListName = "") Then
				strSiteListName = Split(strServerName, ":", 2)(0) & "(" & intOrder & ")"
			Else
				strSiteListName = strSiteListName & ";" & Split(strServerName, ":", 2)(0) & "(" & intOrder & ")"
			End If
			If (strSiteListIP = "") Then
				strSiteListIP = Split(strServerIP, ":", 2)(0) & "(" & intOrder & ")"
			Else
				strSiteListIP = strSiteListIP & ";" & Split(strServerIP, ":", 2)(0) & "(" & intOrder & ")"
			End If
			If (strSiteListPort = "") Then
				strSiteListPort = intSecurePort & "(" & intOrder & ")"
			Else
				strSiteListPort = strSiteListPort & ";" & intSecurePort & "(" & intOrder & ")"
			End If
			rsSiteList.MoveNext
		Wend
	End If
	'
	' Pull "Configuration Manager" data from SMSCFG.ini file
	'
	strSMSConfigFile = "c:\Windows\SMSCFG.ini"
	strSiteSMSConfigFile = "\\" & strPassedParameter & "\" & Replace(strSMSConfigFile, ":", "$")
	If (g_objFSO.FileExists(strSiteSMSConfigFile)) Then
		Set objSiteSMSConfigFile = g_objFSO.OpenTextFile(strSiteSMSConfigFile)
		While Not objSiteSMSConfigFile.AtEndOfStream
			strLine = UCase(Trim(objSiteSMSConfigFile.ReadLine))
			If (InStr(1, strLine, "SMS Unique Identifier=GUID:", vbTextCompare) > 0) Then
				strSMSUID = Split(strLine, "SMS Unique Identifier=GUID:", 2, vbTextCompare)(1)
			End If
			If (InStr(1, strLine, "Previous SMSUID=GUID:", vbTextCompare) > 0) Then
				strPreviousSMSUID = Split(strLine, "Previous SMSUID=GUID:", 2, vbTextCompare)(1)
			End If
			If (InStr(1, strLine, "Last SMSUID Change Date=", vbTextCompare) > 0) Then
				dtLastChangedSMSUID = Split(strLine, "Last SMSUID Change Date=", 2, vbTextCompare)(1)
				dtLastChangedSMSUID = g_objFunctions.MassageTimestamp(dtLastChangedSMSUID)
			End If
			If (InStr(1, strLine, "SID=", vbTextCompare) > 0) Then
				strMachineSID = Split(strLine, "SID=", 2, vbTextCompare)(1)
			End If
			If (InStr(1, strLine, "SMS SMBIOS Serial Number Identifier=", vbTextCompare) > 0) Then
				strSMBIOSSerialNumber = Split(strLine, "SMS SMBIOS Serial Number Identifier=", 2, vbTextCompare)(1)
			End If
			If (InStr(1, strLine, "SMS Hardware Identifier=", vbTextCompare) > 0) Then
				strHardwareIdentifier = Split(strLine, "SMS Hardware Identifier=", 2, vbTextCompare)(1)
			End If
			If (InStr(1, strLine, "SMS Hardware Identifier 2=", vbTextCompare) > 0) Then
				strHardwareIdentifier2 = Split(strLine, "SMS Hardware Identifier 2=", 2, vbTextCompare)(1)
			End If
			If (InStr(1, strLine, "Last Version=", vbTextCompare) > 0) Then
				strLastVersion = Split(strLine, "Last Version=", 2, vbTextCompare)(1)
			End If
		Wend
	End If
	'
	' Get values that are either WMI or registry
	'
	If ((blnWMIGoodToGo) Or (blnWMIRegGoodToGo) Or (blnRemoteRegGoodToGo)) Then
		If (blnWMIGoodToGo) Then
			strSQLQuery = "SELECT Name FROM Win32_ComputerSystem"
			Call g_objFunctions.ExecWMI(objRemoteWMIServer, intErrNumber, strErrDescription, colWMI, strSQLQuery, intFlag, Null)
			If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
				'
				' Run the loop even though there should only be one entry
				'
				For Each objWMI In colWMI
					'
					' The DNSHostName field should contain the Host Name of the computer.  If the computer's host
					' name is longer than 15 characters then this value will be different than the computer's
					' ComputerName.  The Domain field should contain the DNSName of the domain (i.e. att.yahoo.com).
					'
					strComputerName = g_objFunctions.VerifyAndLoad(objWMI.Name, vbString)
				Next
			End If
		End If
		If (strComputerName = "") Then
			'
			' Didn't get the ComputerName via WMI or WMI wasn't available
			'
			If ((blnWMIRegGoodToGo) Or (blnRemoteRegGoodToGo)) Then
				'
				' WMI didn't get it for some reason - try the registry
				'
				strRegistryHive = "HKLM"
				strRegistryKey = "SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName"
				strSearchEntry = "ComputerName"
				Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
																blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
																blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
																strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
				If (blnKeyOrValueExists) Then
					strComputerName = valRegValue
				End If
			End If
		End If

		strWindowsDirectory = ""
		strSystemDirectory = ""
		If (blnWMIGoodToGo) Then
			strSQLQuery = "SELECT WindowsDirectory FROM Win32_OperatingSystem"
			Call g_objFunctions.ExecWMI(objRemoteWMIServer, intErrNumber, strErrDescription, colWMI, strSQLQuery, intFlag, Null)
			If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
				For Each objWMI In colWMI
					strWindowsDirectory = g_objFunctions.VerifyAndLoad(objWMI.WindowsDirectory, vbString)		' C:|WINDOWS
					strSystemDirectory = strWindowsDirectory & "\System32"
				Next
			End If
		End If
		If (strWindowsDirectory = "") Then
			If ((blnWMIRegGoodToGo) Or (blnRemoteRegGoodToGo)) Then
				strRegistryHive = "HKLM"
				strRegistryKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
				strSearchEntry = "SystemRoot"
				Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
																blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
																blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
																strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
				If (blnKeyOrValueExists) Then
					strWindowsDirectory = valRegValue
					strSystemDirectory = strWindowsDirectory & "\System32"
				End If
			End If
		End If
	End If
	'
	' Get Registry values
	'
	If ((blnWMIRegGoodToGo) Or (blnRemoteRegGoodToGo)) Then
		'
		' Setup Registry processing
		'
		strRegistryHive = "HKLM"
		strRegistryKey = "SOFTWARE\Microsoft\SMS\Mobile Client"
		strSearchEntry = "AssignedSiteCode"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			strManuallyAssignedSiteCode = valRegValue
		End If
		strSearchEntry = "GPRequestedSiteAssignmentCode"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			strGPOAssignedSiteCode = valRegValue
		End If
		strSearchEntry = "ProductVersion"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			strSCCMVersion = valRegValue
		End If
		strRegistryKey = "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
		strSearchEntry = "WUServer"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			strWindowsUpdateServer = valRegValue
		End If

		strRegistryKey = "SOFTWARE\Microsoft\OLE"
		strSearchEntry = "EnableDCOM"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			If (valRegValue <> "N") Then
				blnEnableDCOM = True
			End If
		End If
' 		strNameSpaceForConnection = "root\ccm"
' 		strNameSpaceToValidate = "invagt"
' 		strNameSpace = "root\ccm\invagt"
' 		blnNamespaceInstalled = g_objFunctions.IsNamespaceInstalled(strWMIConnectedWithThis, strNameSpaceForConnection, _
' 																		strNameSpaceToValidate, "", "", g_objLogAndTraceErrors)
' 		If (blnNamespaceInstalled) Then
' 			intRetVal = g_objFunctions.CreateServerConnection(strWMIConnectedWithThis, objSMS, intErrNumber, strErrDescription, _
' 																strError, strNameSpace, "", "", g_objLogAndTraceErrors)
' 			If (intRetVal = 0) Then
' 				strSQLQuery = "SELECT ADSiteName FROM CCM_ADSiteInfo"
' 				Call g_objFunctions.ExecWMI(objSMS, intErrNumber, strErrDescription, colWMI, strSQLQuery, wbemFlagReturnWhenComplete, Null)
' 				If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
' 					For Each objWMI In colWMI
' 						WScript.Echo objWMI.AdSiteName
' 						strADSiteName = g_objFunctions.VerifyAndLoad(objWMI.ADSiteName, vbString)
' 					Next
' 				Else
' 					strADSiteName = "Unknown"
' 				End If
' 			End If
' 		End If
		strRegistryKey = "SYSTEM\CurrentControlSet\services\Netlogon\Parameters"
		strSearchEntry = "DynamicSiteName"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			strADSiteName = valRegValue
		End If

		strRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation"
		strSearchEntry = "Model"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			strSDCVersion = valRegValue
		End If
		'
		' Check to see if a reboot is pending.  There are 3 registry values that need to be looked at.  If any one of
		' them is set then no further processing is required.
		'
		strRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing"
		strSearchEntry = "RebootPending"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			blnPendingReboot = True
		End If
		If (blnPendingReboot = False) Then
			strRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"
			strSearchEntry = "RebootRequired"
			Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
															blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
															blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
															strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
			If (blnKeyOrValueExists) Then
				blnPendingReboot = True
			End If
		End If
		If (blnPendingReboot = False) Then
			strRegistryKey = "SYSTEM\CurrentControlSet\Control\Session Manager"
			strSearchEntry = "PendingFileRenameOperations"
			Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
															blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
															blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
															strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
			If (blnKeyOrValueExists) Then
				blnPendingReboot = True
			End If
		End If

		If (strIPAddress <> "") Then
			strRegistryHive = "HKLM"
			strRegistryKey = "SYSTEM\CurrentControlSet\services\Tcpip\Parameters\Interfaces"
			Call g_objRegistryProcessing.EnumRegistryKeys(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, blnRemoteRegGoodToGo, _
																strRegistryHive, strRegistryKey, blnIs64BitMachine, strOSVersion, _
																blnKeyOrValueExists, rsKeys, g_objLogAndTrace, g_objLogAndTraceErrors, _
																g_objLogAndTraceLoadRS)
			If ((blnKeyOrValueExists) And (rsKeys.RecordCount > 0)) Then
				'
				' Save a copy of rsKeys to use for processing.  The call to EnumRegistryKeys deletes
				' any existing records in the recordset during processing.
				'
				blnISaySo = True
				If (Not rsKeys.BOF) Then
					rsKeys.MoveFirst
				End If
				While Not rsKeys.EOF And blnISaySo = True
					strRegistryHive = rsKeys("Hive")
					strRegistryKey = rsKeys("Key")
					strSubkey = rsKeys("Subkey")
					strNewRegistryKey = strRegistryKey & "\" & strSubkey
					blnIsWow6432Node = rsKeys("Wow6432Node")
					Call g_objRegistryProcessing.EnumRegistryEntriesByKey(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
																			blnRemoteRegGoodToGo, strRegistryHive, strNewRegistryKey, _
																			blnIs64BitMachine, blnIsWow6432Node, strOSVersion, _
																			blnKeyOrValueExists, rsEntries, g_objLogAndTrace, g_objLogAndTraceErrors, _
																			g_objLogAndTraceLoadRS)
					If ((blnKeyOrValueExists) And (rsEntries.RecordCount > 0)) Then
						blnEntryExists = ValidateEntryExists(rsEntries, "EnableDHCP", valRegValue)
						blnDHCPEnabled = False
						If (blnEntryExists) Then
							If ((Not IsNull(valRegValue)) And (Not IsEmpty(valRegValue)) And _
								(valRegValue <> "") And (valRegValue <> " ")) Then
								blnDHCPEnabled = valRegValue
							End If
						End If
						If (blnDHCPEnabled) Then
							blnEntryExists = ValidateEntryExists(rsEntries, "DhcpIPAddress", valRegValue)
							strIPAddressFromRegistry = ""
							If (blnEntryExists) Then
								If ((Not IsNull(valRegValue)) And (Not IsEmpty(valRegValue)) And _
									(valRegValue <> "") And (valRegValue <> " ")) Then
									strIPAddressFromRegistry = valRegValue
								End If
							End If
							If (InStr(strIPAddressFromRegistry, strIPAddress) > 0) Then
								'
								' Match
								'
								blnEntryExists = ValidateEntryExists(rsEntries, "DHCPNameServer", valRegValue)
								If (blnEntryExists) Then
									If ((Not IsNull(valRegValue)) And (Not IsEmpty(valRegValue)) And _
										(valRegValue <> "") And (valRegValue <> " ")) Then
										strNameServers = Replace(Replace(valRegValue, ",", ";"), " ", ";")
										blnISaySo = False
									End If
								End If
							End If
						Else
							blnEntryExists = ValidateEntryExists(rsEntries, "IPAddress", valRegValue)
							strIPAddressFromRegistry = ""
							If (blnEntryExists) Then
								If ((Not IsNull(valRegValue)) And (Not IsEmpty(valRegValue)) And _
									(valRegValue <> "") And (valRegValue <> " ")) Then
									strIPAddressFromRegistry = valRegValue
								End If
							End If
							If (InStr(strIPAddressFromRegistry, strIPAddress) > 0) Then
								'
								' Match
								'
								blnEntryExists = ValidateEntryExists(rsEntries, "NameServer", valRegValue)
								If (blnEntryExists) Then
									If ((Not IsNull(valRegValue)) And (Not IsEmpty(valRegValue)) And _
										(valRegValue <> "") And (valRegValue <> " ")) Then
										strNameServers = Replace(Replace(valRegValue, ",", ";"), " ", ";")
										blnISaySo = False
									End If
								End If
							End If
						End If
					End If
					rsKeys.MoveNext
				Wend
			End If
		End If
		'
		' HBSS Registry settings
		'		
		strRegistryKey = "SOFTWARE\McAfee\AVEngine"
		strSearchEntry = "AVDatDate"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			If ((IsNull(valRegValue)) Or (IsEmpty(valRegValue)) Or (valRegValue = "") Or (valRegValue = " ")) Then
				dtAVDatDate = DEFAULT_DATE
			Else
				dtAVDatDate = valRegValue
			End If
		End If

		strRegistryKey = "SOFTWARE\Network Associates\ePolicy Orchestrator\Agent"
		strSearchEntry = "LastASCTime"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			dtLastASCTime = DateAdd("s", valRegValue, DEFAULT_DATE)
			dtLastASCTime = g_objFunctions.MassageTimestamp(dtLastASCTime)
		End If

		strSearchEntry = "PropsVersion"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			dtPropsVersionDate = g_objFunctions.WMIDateStringToDate(valRegValue)
		End If

		strSearchEntry = "AgentWakeUpPort"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			intAgentWakeUpPort = valRegValue
		End If

		strSearchEntry = "ePOServerList"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			g_objFunctions.DeleteAllRecordsetRows(rsSiteList)
			If (InStr(valRegValue, ";") > 0) Then
				arrServerList = Split(valRegValue, ";")
				For intCount = 0 To UBound(arrServerList)
					arrSplit = Split(arrServerList(intCount), "|")
					strServerName = arrSplit(0)
					strServerIP = arrSplit(1)
					intSecurePort = CInt(arrSplit(2))
					rsSiteList.AddNew
					rsSiteList("Order") = intCount + 1
					rsSiteList("ServerName") = strServerName
					rsSiteList("ServerIP") = strServerIP
					rsSiteList("SecurePort") = intSecurePort
					rsSiteList.Update
				Next
			Else
				arrSplit = Split(valRegValue, "|")
				strServerName = arrSplit(0)
				strServerIP = arrSplit(1)
				intSecurePort = CInt(arrSplit(2))
				rsSiteList.AddNew
				rsSiteList("Order") = 1
				rsSiteList("ServerName") = strServerName
				rsSiteList("ServerIP") = strServerIP
				rsSiteList("SecurePort") = intSecurePort
				rsSiteList.Update
			End If
		End If
		If (rsSiteList.RecordCount > 0) Then
			rsSiteList.MoveFirst
			While Not rsSiteList.EOF
				intOrder = rsSiteList("Order")
				strServerName = rsSiteList("ServerName")
				strServerIP = rsSiteList("ServerIP")
				intSecurePort = rsSiteList("SecurePort")
				If (strEPORegistryName = "") Then
					strEPORegistryName = Split(strServerName, ":", 2)(0) & "(" & intOrder & ")"
				Else
					strEPORegistryName = strEPORegistryName & ";" & Split(strServerName, ":", 2)(0) & "(" & intOrder & ")"
				End If
				If (strEPORegistryIP = "") Then
					strEPORegistryIP = Split(strServerIP, ":", 2)(0) & "(" & intOrder & ")"
				Else
					strEPORegistryIP = strEPORegistryIP & ";" & Split(strServerIP, ":", 2)(0) & "(" & intOrder & ")"
				End If
				If (strEPORegistryPort = "") Then
					strEPORegistryPort = intSecurePort & "(" & intOrder & ")"
				Else
					strEPORegistryPort = strEPORegistryPort & ";" & intSecurePort & "(" & intOrder & ")"
				End If
				rsSiteList.MoveNext
			Wend
		End If

		strRegistryKey = "SOFTWARE\Network Associates\ePolicy Orchestrator\Agent"
		strSearchEntry = "AgentGUID"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			strAgentGUID = Replace(Replace(valRegValue, "{", ""), "}", "")
		End If

		strRegistryKey = "SOFTWARE\Network Associates\ePolicy Orchestrator\Application Plugins\EPOAGENT3000"
		strSearchEntry = "Version"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			strMAVersion = valRegValue
		End If
		
		strRegistryKey = "SOFTWARE\McAfee\HIP\Config\Settings"
		strSearchEntry = "FW_Enabled"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			blnFWEnabled = valRegValue
		End If
		
		strRegistryKey = "SOFTWARE\McAfee\DesktopProtection"
		strSearchEntry = "szProductVer"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			strVSEVersion = valRegValue
		End If

		strRegistryKey = "SOFTWARE\McAfee\HIP"
		strSearchEntry = "Version"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			strHIPSVersion = valRegValue
		End If

		strRegistryKey = "SOFTWARE\McAfee\DLP\Agent"
		strSearchEntry = "AgentVersion"
		Call g_objRegistryProcessing.GetRegistryEntry(objRemoteRegServer, strRemoteRegConnectedWithThis, blnWMIRegGoodToGo, _
														blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strSearchEntry, _
														blnIs64BitMachine, strOSVersion, valRegValue, blnKeyOrValueExists, _
														strKeyType, g_objLogAndTrace, g_objLogAndTraceErrors, g_objLogAndTraceLoadRS)
		If (blnKeyOrValueExists) Then
			strDLPVersion = valRegValue
		End If
	End If
	'
	' Last SMS patch received when?
	'
	strSCCMCacheFolder = ""
	If (strSCCMVersion <> "") Then
		If (Mid(strSCCMVersion, 1, 1) = "4") Then
			'
			' SCCM 2007
			'
			If (intOSBits = 64) Then
				strSCCMCacheFolder = strWindowsDirectory & "\SysWOW64\CCM\Cache"
			Else
				strSCCMCacheFolder = strSystemDirectory & "\CCM\Cache"
			End If
		ElseIf (Mid(strSCCMVersion, 1, 1) = "5") Then
			'
			' SCCM 2012
			'
			strSCCMCacheFolder = strWindowsDirectory & "\CCMCache"
		Else
			strSCCMCacheFolder = ""
		End If
	End If
	dtMostRecent = DEFAULT_DATE
	strMostRecent = ""

	If (strSCCMCacheFolder <> "") Then
		If ((g_blnProcessingLocal = False) And (blnWMIGoodToGo)) Then
			'
			' Does the directory exist?
			'
			If (g_objFunctions.WMIFolderExists(objRemoteWMIServer, strSCCMCacheFolder, wbemFlagReturnWhenComplete)) Then
				'
				' The folder exists on the remote computer - process each folder to find the most recently updated one
				'
				Set colFolders = objRemoteWMIServer.ExecQuery("ASSOCIATORS OF {Win32_Directory.Name='" & strSCCMCacheFolder & "'} " _
													& "WHERE AssocClass = Win32_Subdirectory " _
													& "ResultRole = PartComponent",, intFlag)
				For Each objSubFolder In colFolders
					dtCreationDate = g_objFunctions.VerifyAndLoad(objSubFolder.CreationDate, vbDate)
					'
					' Find the most recent based on creation date
					'
					If (CDate(dtMostRecent) < CDate(dtCreationDate)) Then
						dtMostRecent = dtCreationDate
						strMostRecent = objSubFolder.Name
					End If
				Next
			End If
		Else
			'
			' Does the directory exist?
			'
			If (g_objFSO.FolderExists(strSCCMCacheFolder)) Then
				'
				' Local processing or remote processing not using WMI (FSO)
				'
				' The folder exists on the remote computer - process each folder to find the most recently updated one
				'
				If (g_blnProcessingLocal) Then
					strFolderName = strSCCMCacheFolder
				Else
					strFolderName = "\\" & strPassedParameter & "\" & Replace(strSCCMCacheFolder, ":\", "$\")
				End If
				On Error Resume Next
				Set colFolders = g_objFSO.GetFolder(strFolderName)
				intErrNumber = Err.Number
				strErrDescription = Err.Description
				On Error GoTo 0
				If (intErrNumber = 0) Then
					Set colSubFolders = colFolders.SubFolders
					For Each objSubFolder in colSubFolders
						dtCreationDate = g_objFunctions.VerifyAndLoad(objSubFolder.DateCreated, vbDate)
						'
						' Find the most recent based on creation date
						'
						If (CDate(dtMostRecent) < CDate(dtCreationDate)) Then
							dtMostRecent = dtCreationDate
							strMostRecent = objSubFolder.Path
						End If
					Next
				End If
			End If
		End If
	End If

	If (strSystemDirectory <> "") Then
		'
		' Check for trailing \
		'
		If (Right(strSystemDirectory, 1) <> "\") Then
			strSystemDirectory = strSystemDirectory & "\"
		End If
		'
		' The CIM_Datafile call requires slashes to become double slashes
		'
		strSystemDirectoryForConnect = Replace(strSystemDirectory, "\", "\\")
		strSystemDirectoryForConnect = strSystemDirectoryForConnect & "wuaueng.dll"
		'
		' Get file information
		'
		strSQLQuery = "SELECT Version FROM CIM_DataFile WHERE Name = '" & Replace(strSystemDirectoryForConnect, "'", "''") & "'"
		Call g_objFunctions.ExecWMI(objRemoteWMIServer, intErrNumber, strErrDescription, colFiles, strSQLQuery, intFlag, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colFiles))="SWBEMOBJECTSET")) Then
			If (colFiles.Count > 0) Then
				'
				' The file exists
				'
				For Each objFile In colFiles
					strWUAVersion = objFile.Version
				Next
			End If
		End If
	End If
	'
	' Connect to the Advanced Client Namespace
	'
	strNameSpaceForConnection = "root"
	strNameSpaceToValidate = "ccm"
	strNameSpace = "root\ccm"
	blnNamespaceInstalled = g_objFunctions.IsNamespaceInstalled(strWMIConnectedWithThis, strNameSpaceForConnection, _
																	strNameSpaceToValidate, "", "", g_objLogAndTraceErrors)
	If (blnNamespaceInstalled) Then
		intRetVal = g_objFunctions.CreateServerConnection(strWMIConnectedWithThis, objSMS, intErrNumber, strErrDescription, _
															strError, strNameSpace, "", "", g_objLogAndTraceErrors)
		If (intRetVal = 0) Then
			strSQLQuery = "SELECT Name,CurrentManagementPoint FROM SMS_Authority"
			Call g_objFunctions.ExecWMI(objSMS, intErrNumber, strErrDescription, colWMI, strSQLQuery, wbemFlagReturnWhenComplete, Null)
			If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
				For Each objWMI In colWMI
					strClientSite = Replace(g_objFunctions.VerifyAndLoad(objWMI.Name, vbString), "SMS:", "")
					If (objWMI.CurrentManagementPoint <> "") Then
						strCurrentMP = g_objFunctions.VerifyAndLoad(objWMI.CurrentManagementPoint, vbString)
					End If
				Next
			Else
				strClientSite = "Unknown"
				strCurrentMP = "SMS_Authority Query (root\ccm) failed with error: " & intErrNumber & "  Desc: " & strErrDescription
			End If
		End If
	End If
	'
	' Get Running Advertisements
	'
	strNameSpaceForConnection = "root\ccm"
	strNameSpaceToValidate = "SoftMgmtAgent"
	strNameSpace = "root\ccm\SoftMgmtAgent"
	blnNamespaceInstalled = g_objFunctions.IsNamespaceInstalled(strWMIConnectedWithThis, strNameSpaceForConnection, _
																	strNameSpaceToValidate, "", "", g_objLogAndTraceErrors)
	If (blnNamespaceInstalled) Then
		intRetVal = g_objFunctions.CreateServerConnection(strWMIConnectedWithThis, objSMS, intErrNumber, strErrDescription, _
															strError, strNameSpace, "", "", g_objLogAndTraceErrors)
		If (intRetVal = 0) Then
			strSQLQuery = "SELECT * FROM CCM_ExecutionRequestEx"
			Call g_objFunctions.ExecWMI(objSMS, intErrNumber, strErrDescription, colWMI, strSQLQuery, intFlag, Null)
			If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
				'
				' List all running advertisements on client
				'
				intRunningAdvertisements = 0
				For Each objWMI In colWMI
					intRunningAdvertisements = intRunningAdvertisements + 1
				Next
			End If
		End If
	End If
	Call g_objFunctions.DeleteAllRecordsetRows(g_rsCertificates)
	Call GetCertificateSettingsSMS(objRemoteRegServer, blnWMIRegGoodToGo, blnRemoteRegGoodToGo, strLocalAccountsConnectedWithThis, _
									blnIs64BitMachine, strOSVersion)
	If (g_rsCertificates.RecordCount > 0) Then
		If (Not g_rsCertificates.BOF) Then
			g_rsCertificates.MoveFirst
		End If
		While Not g_rsCertificates.EOF
			If (UCase(g_rsCertificates("Type")) = "ENCRYPTION") Then
				strEncryptionSubject = g_rsCertificates("Subject")
				dtEncryptionNotBefore = g_rsCertificates("NotBefore")
				dtEncryptionNotAfter = g_rsCertificates("NotAfter")
				intEncryptionCertNumber = g_rsCertificates("CertNumber")
				strEncryptionRegKey = g_rsCertificates("RegistryKey")
				If ((intEncryptionCertNumber <> -1) Or (strEncryptionRegKey <> "FFFFFFFF")) Then
					If (InStr(1, strEncryptionSubject, strComputerName, vbTextCompare) > 0) Then
						blnEncryptionSubjectMatch = True
					End If
				End If
			Else
				strSigningSubject = g_rsCertificates("Subject")
				dtSigningNotBefore = g_rsCertificates("NotBefore")
				dtSigningNotAfter = g_rsCertificates("NotAfter")
				intSigningCertNumber = g_rsCertificates("CertNumber")
				strSigningRegKey = g_rsCertificates("RegistryKey")
				If ((intSigningCertNumber <> -1) Or (strSigningRegKey <> "FFFFFFFF")) Then
					If (InStr(1, strSigningSubject, strComputerName, vbTextCompare) > 0) Then
						blnSigningSubjectMatch = True
					End If
				End If
			End If
			g_rsCertificates.MoveNext
		Wend
	End If
	dtThirtyDaysAgo = DateAdd("d", -30, Now())
	dtSevenDaysAgo = DateAdd("d", -7, Now())
	'
	' SCCM Healthy?
	'
	If ((dtThirtyDaysAgo <= dtLastBootupTime) And _
		(intPercentFreeSpace > 15) And _
		(strManuallyAssignedSiteCode <> "") And _
		(strGPOAssignedSiteCode <> "") And _
		(blnEncryptionSubjectMatch <> 0) And _
		(blnSigningSubjectMatch <> 0) And _
		(dtEncryptionNotBefore < Now()) And _
		(Now() < dtEncryptionNotAfter) And _
		(dtSigningNotBefore < Now()) And _
		(Now() < dtSigningNotAfter) And _
		(UCase(strWUAUState) = "RUNNING") And (UCase(strWUAUStartMode) = "AUTO") And _
		(UCase(strSCCMState) = "RUNNING") And (UCase(strSCCMStartMode) = "AUTO") And _
		((UCase(strBITSStartMode) = "AUTO") Or (UCase(strBITSStartMode) = "MANUAL")) And _
		(UCase(strWinMgmtState) = "RUNNING") And (UCase(strWinMgmtStartMode) = "AUTO") And _
		(((UCase(strGPClientState) = "RUNNING") Or (UCase(strGPClientState) = "NA")) And ((UCase(strGPClientStartMode) = "AUTO") Or (UCase(strGPClientStartMode) = "NA"))) And _
		(UCase(strLanmanServerState) = "RUNNING") And (UCase(strLanmanServerStartMode) = "AUTO") And _
		(UCase(strRPCSSState) = "RUNNING") And (UCase(strRPCSSStartMode) = "AUTO") And _
		(UCase(strRemoteRegistryState) = "RUNNING") And (UCase(strRemoteRegistryStartMode) = "AUTO") And _
		(UCase(strSMSTSMGRStartMode) = "MANUAL") And _
		(blnWMIGoodToGo) And _
		(blnWMIRegGoodToGo) And _
		(blnPendingReboot = False) And _
		(blnEnableDCOM)) Then
		blnSCCMHealthy = True
	End If
	'
	' HBSS Healthy?
	'
	If ((dtThirtyDaysAgo <= dtLastBootupTime) And _
		(UCase(strMcFrameworkState) = "RUNNING") And (UCase(strMcFrameworkStartMode) = "AUTO") And _
		(UCase(strMcShieldState) = "RUNNING") And (UCase(strMcShieldStartMode) = "AUTO") And _
		(dtSevenDaysAgo <= dtAVDatDate) And _
		(dtSevenDaysAgo <= dtCatalogVersionDate) And _
		(blnWMIGoodToGo) And _
		(blnWMIRegGoodToGo) And _
		(blnRemoteRegGoodToGo) And _
		(blnPendingReboot = False)) Then
		blnHBSSHealthy = True
	End If
	'
	' Create the recordset entry for the data collected
	'
	g_rsClientConfiguration.AddNew
	g_rsClientConfiguration("ComputerFQDN") = LCase(strComputerFQDN)
	g_rsClientConfiguration("ComputerName") = UCase(strComputerName)
	g_rsClientConfiguration("IPAddress") = strIPAddress
	g_rsClientConfiguration("OSName") = strOSName
	g_rsClientConfiguration("OSBuildNumber") = strOSBuildNumber
	g_rsClientConfiguration("LastBootupTime") = dtLastBootupTime
	g_rsClientConfiguration("OSBuildType") = strOSBuildType
	g_rsClientConfiguration("OSType") = intOSType
	g_rsClientConfiguration("OSTypeText") = strOSType
	g_rsClientConfiguration("ProductType") = intProductType
	g_rsClientConfiguration("ProductTypeText") = strProductType
	g_rsClientConfiguration("OSVersion") = strOSVersion
	g_rsClientConfiguration("AddressWidth") = strOSBits
	g_rsClientConfiguration("PercentFreeC") = intPercentFreeSpace
	g_rsClientConfiguration("ManuallyAssignedSiteCode") = strManuallyAssignedSiteCode
	g_rsClientConfiguration("GPOAssignedSiteCode") = strGPOAssignedSiteCode
	g_rsClientConfiguration("EncryptionSubject") = strEncryptionSubject
	g_rsClientConfiguration("EncryptionSubjectMatch") = blnEncryptionSubjectMatch
	g_rsClientConfiguration("EncryptionCertNumber") = intEncryptionCertNumber
	g_rsClientConfiguration("EncryptionRegistryKey") = strEncryptionRegKey
	g_rsClientConfiguration("EncryptionNotBefore") = dtEncryptionNotBefore
	g_rsClientConfiguration("EncryptionNotAfter")  = dtEncryptionNotAfter
	g_rsClientConfiguration("SigningSubject") = strSigningSubject
	g_rsClientConfiguration("SigningSubjectMatch") = blnSigningSubjectMatch
	g_rsClientConfiguration("SigningCertNumber") = intSigningCertNumber
	g_rsClientConfiguration("SigningRegistryKey") = strSigningRegKey
	g_rsClientConfiguration("SigningNotBefore") = dtSigningNotBefore
	g_rsClientConfiguration("SigningNotAfter")  = dtSigningNotAfter
	g_rsClientConfiguration("WUAUState") = strWUAUState
	g_rsClientConfiguration("WUAUStartMode") = strWUAUStartMode
	g_rsClientConfiguration("CCMEXECState") = strSCCMState
	g_rsClientConfiguration("CCMEXECStartMode") = strSCCMStartMode
	g_rsClientConfiguration("BITSState") = strBITSState
	g_rsClientConfiguration("BITSStartMode") = strBITSStartMode
	g_rsClientConfiguration("WinMgmtState") = strWinMgmtState
	g_rsClientConfiguration("WinMgmtStartMode") = strWinMgmtStartMode
	g_rsClientConfiguration("GPClientState") = strGPClientState
	g_rsClientConfiguration("GPClientStartMode") = strGPClientStartMode
	g_rsClientConfiguration("SCCMVersion") = strSCCMVersion
	g_rsClientConfiguration("WMIGoodToGo") = blnWMIGoodToGo
	g_rsClientConfiguration("WMIRegGoodToGo") = blnWMIRegGoodToGo
	g_rsClientConfiguration("RemoteRegGoodToGo") = blnRemoteRegGoodToGo
	g_rsClientConfiguration("WindowsUpdateServer") = strWindowsUpdateServer
	g_rsClientConfiguration("NameServers") = strNameServers
	g_rsClientConfiguration("LastLoggedOnUser") = strLastLoggedOnUser
	g_rsClientConfiguration("PendingReboot") = blnPendingReboot
	g_rsClientConfiguration("SMSUID") = strSMSUID
	g_rsClientConfiguration("PreviousSMSUID") = strPreviousSMSUID
	g_rsClientConfiguration("LastChangedSMSUID") = dtLastChangedSMSUID
	g_rsClientConfiguration("ClientSite") = strClientSite
	g_rsClientConfiguration("CurrentMP") = strCurrentMP
	g_rsClientConfiguration("ADSiteName") = strADSiteName
	g_rsClientConfiguration("SDCVersion") = strSDCVersion
	g_rsClientConfiguration("WUAVersion") = strWUAVersion
	g_rsClientConfiguration("CCMSetupState") = strCCMSetupState
	g_rsClientConfiguration("CCMSetupStartMode") = strCCMSetupStartMode
	g_rsClientConfiguration("LanmanServerState") = strLanmanServerState
	g_rsClientConfiguration("LanmanServerStartMode") = strLanmanServerStartMode
	g_rsClientConfiguration("RPCSSState") = strRPCSSState
	g_rsClientConfiguration("RPCSSStartMode") = strRPCSSStartMode
	g_rsClientConfiguration("SMSTaskSequenceAgentState") = strSMSTSMGRState
	g_rsClientConfiguration("SMSTaskSequenceAgentStartMode") = strSMSTSMGRStartMode
	g_rsClientConfiguration("RemoteRegistryState") = strRemoteRegistryState
	g_rsClientConfiguration("RemoteRegistryStartMode") = strRemoteRegistryStartMode
	g_rsClientConfiguration("EnableDCOM") = blnEnableDCOM
	g_rsClientConfiguration("MostRecentSMSFolder") = strMostRecent
	g_rsClientConfiguration("MostRecentSMSFolderDate") = dtMostRecent
	g_rsClientConfiguration("SCCMHealthy") = blnSCCMHealthy
	g_rsClientConfiguration("RunningAdvertisements") = intRunningAdvertisements
	'
	' HBSS settings
	'
	g_rsClientConfiguration("AVDatDate") = dtAVDatDate
	g_rsClientConfiguration("CatalogVersionDate") = dtCatalogVersionDate
	g_rsClientConfiguration("SiteListName") = strSiteListName
	g_rsClientConfiguration("SiteListIP") = strSiteListIP
	g_rsClientConfiguration("SiteListPort") = strSiteListPort
	g_rsClientConfiguration("EPORegistryName") = strEPORegistryName
	g_rsClientConfiguration("EPORegistryIP") = strEPORegistryIP
	g_rsClientConfiguration("EPORegistryPort") = strEPORegistryPort
	g_rsClientConfiguration("AgentGUID") = strAgentGUID
	g_rsClientConfiguration("McFrameworkState") = strMcFrameworkState
	g_rsClientConfiguration("McFrameworkStartMode") = strMcFrameworkStartMode
	g_rsClientConfiguration("McShieldState") = strMcShieldState
	g_rsClientConfiguration("McShieldStartMode") = strMcShieldStartMode
	g_rsClientConfiguration("LastASCTime") = dtLastASCTime
	g_rsClientConfiguration("PropsVersionDate") = dtPropsVersionDate
	g_rsClientConfiguration("AgentWakeUpPort") = intAgentWakeUpPort
	g_rsClientConfiguration("MAVersion") = strMAVersion
	g_rsClientConfiguration("FWEnabled") = blnFWEnabled
	g_rsClientConfiguration("VSEVersion") = strVSEVersion
	g_rsClientConfiguration("HIPSVersion") = strHIPSVersion
	g_rsClientConfiguration("DLPVersion") = strDLPVersion
	g_rsClientConfiguration("HBSSHealthy") = blnHBSSHealthy
	g_rsClientConfiguration.Update
	'
	' Cleanup
	'
	Set colWMI = Nothing
	Set objValidationErrorsFile = Nothing
	Set rsKeys = Nothing
	Set rsEntries = Nothing
	Set objFrameworkINIFile = Nothing
	Set rsSiteList = Nothing
	Set xmlDoc = Nothing
	Set colRowNodes = Nothing
	Set colAttributes = Nothing
	GetClientConfiguration = 0

End Function

Function ScanClientConfiguration(ByVal strPassedParameter, ByVal intFlag)
'*****************************************************************************************************************************************
'*  Purpose:				Scans remote machine for configuration information.
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				Mainline
'*  Calls:					GetGMTTimestamp, GetClientConfiguration, CreateNewTableElement, CreateAppendNewElement
'*							CreateSetAndLinkAttribute, CreateNewElement, AppendChild, GetGMTTimestamp, FormatAndSaveXMLFile
'*  Requirements:			None
'*****************************************************************************************************************************************
	Dim dtStartTime, intRetVal, strWMIConnectedWithThis, objRemoteWMIServer, objRemoteRegServer, blnIsClientAlive

	dtStartTime = g_objFunctions.GetGMTTimestamp()
	intRetVal = GetClientConfiguration(strPassedParameter, intFlag, strWMIConnectedWithThis, objRemoteWMIServer, _
										objRemoteRegServer, blnIsClientAlive)
	If (intRetVal = 0) Then
		Call BuildClientHealthXML(strPassedParameter, dtStartTime)
	End If
	'
	' Cleanup
	'
	Set objRemoteWMIServer = Nothing
	Set objRemoteRegServer = Nothing

End Function

Function UpdateService(ByRef objWMIService, ByVal blnWMIGoodToGo, ByVal strPassedParameter, ByVal strService, ByVal strMode, _
							ByRef strState, ByRef objRepairFile)
'*****************************************************************************************************************************************
'*  Purpose:				Does SCCM repair processing on the specified machine.
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				DoRepairProcessing
'*  Calls:					ExecWMI, ExecCmdGeneric, DeleteAllRecordsetRows
'*  Requirements:			None
'*****************************************************************************************************************************************
	Dim blnDoWMIProcessing, strSQLQuery, intErrNumber, strErrDescription, colWMI, objWMI, strCommand, intRetVal, strTemp, blnISaySo, blnStarted

	Const wbemFlagReturnWhenComplete = 0
	blnDoWMIProcessing = False
	If (blnWMIGoodToGo) Then
		strSQLQuery = "SELECT * FROM Win32_Service WHERE Name='" & strService & "'"
		Call g_objFunctions.ExecWMI(objWMIService, intErrNumber, strErrDescription, colWMI, strSQLQuery, wbemFlagReturnWhenComplete, Null)
		If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
			If (IsEmpty(colWMI) = False) Then
				blnDoWMIProcessing = True
			ElseIf (colWMI.Count > 0) Then
				blnDoWMIProcessing = True
			End If
		End If
	End If

	Select Case UCase(strMode)
		Case "STOP"
			If (UCase(strState) = "RUNNING") Then
				objRepairFile.WriteLine("Stopping Service " & strService)
				blnISaySo = True
				While blnISaySo
					If (blnDoWMIProcessing) Then
						For Each objWMI In colWMI
							intRetVal = objWMI.StopService()
						Next
					Else
						strCommand = "sc \\" & strPassedParameter & "stop " & strService
						intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, g_rsGeneric, g_objLogAndTraceExecCmdGeneric)
						g_objFunctions.DeleteAllRecordsetRows(g_rsGeneric)
					End If
					WScript.Sleep 1000
					blnStarted = False
					If (blnDoWMIProcessing) Then
						strSQLQuery = "SELECT * FROM Win32_Service WHERE Name='" & strService & "'"
						Call g_objFunctions.ExecWMI(objWMIService, intErrNumber, strErrDescription, colWMI, strSQLQuery, wbemFlagReturnWhenComplete, Null)
						If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
							For Each objWMI in colWMI
								blnStarted = objWMI.Started
							Next
						End If
					Else
						strCommand = "sc \\" & strPassedParameter & "query " & strService
						intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, g_rsGeneric, g_objLogAndTraceExecCmdGeneric)
						If (Not g_rsGeneric.BOF) Then
							g_rsGeneric.MoveFirst
						End If
						While Not g_rsGeneric.EOF
							strTemp = g_rsGeneric("SavedData")
							If ((InStr(1, strTemp, "STATE") > 0) And (InStr(1, strTemp, "RUNNING"))) Then
								blnStarted = True
							End If
							g_rsGeneric.MoveNext
						Wend
						g_objFunctions.DeleteAllRecordsetRows(g_rsGeneric)
					End If
					If (blnStarted = False) Then
						blnISaySo = False
						strState = "STOPPED"
					End If
				Wend
			End If

		Case "START"
			If (UCase(strState) <> "RUNNING") Then
				objRepairFile.WriteLine("Starting Service " & strService)
				If (blnDoWMIProcessing) Then
					For Each objWMI In colWMI
						intRetVal = objWMI.StartService()
					Next
				Else
					strCommand = "sc \\" & strPassedParameter & "start " & strService
					intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, g_rsGeneric, g_objLogAndTraceExecCmdGeneric)
					g_objFunctions.DeleteAllRecordsetRows(g_rsGeneric)
				End If
			End If
		
		Case "CHANGE"
			If (UCase(strState) = "RUNNING") Then
				objRepairFile.WriteLine("Stopping Service " & strService)
				blnISaySo = True
				While blnISaySo
					If (blnDoWMIProcessing) Then
						For Each objWMI In colWMI
							intRetVal = objWMI.StopService()
						Next
					Else
						strCommand = "sc \\" & strPassedParameter & "stop " & strService
						intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, g_rsGeneric, g_objLogAndTraceExecCmdGeneric)
						g_objFunctions.DeleteAllRecordsetRows(g_rsGeneric)
					End If
					WScript.Sleep 1000
					blnStarted = False
					If (blnDoWMIProcessing) Then
						strSQLQuery = "SELECT * FROM Win32_Service WHERE Name='" & strService & "'"
						Call g_objFunctions.ExecWMI(objWMIService, intErrNumber, strErrDescription, colWMI, strSQLQuery, wbemFlagReturnWhenComplete, Null)
						If ((intErrNumber=0) And (UCase(TypeName(colWMI))="SWBEMOBJECTSET")) Then
							For Each objWMI in colWMI
								blnStarted = objWMI.Started
							Next
						End If
					Else
						strCommand = "sc \\" & strPassedParameter & "query " & strService
						intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, g_rsGeneric, g_objLogAndTraceExecCmdGeneric)
						If (Not g_rsGeneric.BOF) Then
							g_rsGeneric.MoveFirst
						End If
						While Not g_rsGeneric.EOF
							strTemp = g_rsGeneric("SavedData")
							If ((InStr(1, strTemp, "STATE") > 0) And (InStr(1, strTemp, "RUNNING"))) Then
								blnStarted = True
							End If
							g_rsGeneric.MoveNext
						Wend
						g_objFunctions.DeleteAllRecordsetRows(g_rsGeneric)
					End If
					If (blnStarted = False) Then
						blnISaySo = False
						strState = "STOPPED"
					End If
				Wend
			End If
			If ((UCase(strService) = "BITS") Or (UCase(strService) = "SMSTSMGR")) Then
				objRepairFile.WriteLine("Altering Service " & strService & " StartMode to Manual")
			Else
				objRepairFile.WriteLine("Altering Service " & strService & " StartMode to Auto")
			End If
			If (blnDoWMIProcessing) Then
				For Each objWMI In colWMI
'					intRetVal = objWMIService.Change(, , , , "AUTOMATIC")
					If ((UCase(strService) = "BITS") Or (UCase(strService) = "SMSTSMGR")) Then
						objWMI.ChangeStartMode("Manual")
					Else
						objWMI.ChangeStartMode("Automatic")
					End If
				Next
			Else
				If ((UCase(strService) = "BITS") Or (UCase(strService) = "SMSTSMGR")) Then
					strCommand = "sc \\" & strPassedParameter & "config " & strService & " start= manual"
				Else
					strCommand = "sc \\" & strPassedParameter & "config " & strService & " start= auto"
				End If
				intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, g_rsGeneric, g_objLogAndTraceExecCmdGeneric)
				g_objFunctions.DeleteAllRecordsetRows(g_rsGeneric)
			End If
		Case Else
	End Select

End Function

Function GetCertificateInfo(ByVal strCommand, ByVal strMachineToCompare)
'*****************************************************************************************************************************************
'*  Purpose:				Does SCCM repair processing on the specified machine.
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				Mainline
'*  Calls:					ExecCmdGeneric, ParseSMSCertificateSubject, DeleteAllRecordsetRows
'*  Requirements:			None
'*****************************************************************************************************************************************
	Dim strSubject, intRetVal

	strSubject = ""
	intRetVal = g_objFunctions.ExecCmdGeneric(strCommand, g_rsGeneric, g_objLogAndTraceExecCmdGeneric)
	GetCertificateInfo = -1
	If (intRetVal = 0) Then
		'
		' If RecordCount = 2 then the certificate doesn't exist in the CertStore
		'
		If (g_rsGeneric.RecordCount = 2) Then
			GetCertificateInfo = -2
		ElseIf (g_rsGeneric.RecordCount > 2) Then
			Call ParseSMSCertificateSubject(g_rsGeneric, strSubject)
			If (UCase(strSubject) = UCase(strMachineToCompare)) Then
				GetCertificateInfo = 0
			Else
'				WScript.Echo "Compare: " & UCase(strSubject) & vbTab & UCase(strMachineToCompare)
				GetCertificateInfo = -3
			End If
		End If
	End If
	g_objFunctions.DeleteAllRecordsetRows(g_rsGeneric)

End Function

Function DoRepairProcessing(ByVal strPassedParameter, ByVal intFlag)
'*****************************************************************************************************************************************
'*  Purpose:				Does SCCM repair processing on the specified machine.
'*  Arguments supplied:		Look up
'*  Return Value:			0 to indicate success
'*  Called by:				Mainline
'*  Calls:					GetClientConfiguration, BuildDateString, UpdateService, CorrectCertificate, SetRegistryEntry
'*							ScanClientConfiguration
'*  Requirements:			None
'*****************************************************************************************************************************************
	Dim dtStartTime, intRetVal, strWMIConnectedWithThis, objRemoteWMIServer, objRemoteRegServer, blnIsClientAlive, strRepairErrorFile
	Dim objRepairErrorFile, strWUAUState, strWUAUStartMode, strBITSState, strBITSStartMode, strWinMgmtState, strWinMgmtStartMode
	Dim blnEncryptionSubjectMatch, dtEncryptionNotBefore, dtEncryptionNotAfter, blnSigningSubjectMatch, dtSigningNotBefore
	Dim dtSigningNotAfter, strSCCMState, strSCCMStartMode, strGPClientState, strGPClientStartMode, strMcFrameworkState
	Dim strMcFrameworkStartMode, strMcShieldState, strMcShieldStartMode, strLanmanServerState, strLanmanServerStartMode, strRPCSSState
	Dim strRPCSSStartMode, strSMSTSMGRState, strSMSTSMGRStartMode, strRemoteRegistryState, strRemoteRegistryStartMode, blnEnableDCOM
	Dim strManuallyAssignedSiteCode, strEncryptionSubject, intEncryptionCertNumber, strEncryptionRegKey, strSigningSubject
	Dim intSigningCertNumber, strSigningRegKey, blnWMIGoodToGo, blnWMIRegGoodToGo, blnRemoteRegGoodToGo, strOSVersion, strOSBits
	Dim blnIs64BitMachine, strRepairFile, objRepairFile, blnRestartRequired, strRegistryHive, strRegistryKey, strEntryName, strOperation
	Dim blnIsWow6432Node

	dtStartTime = g_objFunctions.GetGMTTimestamp()
	intRetVal = GetClientConfiguration(strPassedParameter, intFlag, strWMIConnectedWithThis, objRemoteWMIServer, _
										objRemoteRegServer, blnIsClientAlive)
	If (intRetVal <> 0) Then
		Exit Function
	End If
	'
	' Check Configuration Items to see if they are ok
	'
	If (blnIsClientAlive) Then
		If (g_rsClientConfiguration.RecordCount > 0) Then
			g_rsClientConfiguration.MoveFirst
			strWUAUState = g_rsClientConfiguration("WUAUState")
			strWUAUStartMode = g_rsClientConfiguration("WUAUStartMode")
			strBITSState = g_rsClientConfiguration("BITSState")
			strBITSStartMode = g_rsClientConfiguration("BITSStartMode")
			strWinMgmtState = g_rsClientConfiguration("WinMgmtState")
			strWinMgmtStartMode = g_rsClientConfiguration("WinMgmtStartMode")
			blnEncryptionSubjectMatch = g_rsClientConfiguration("EncryptionSubjectMatch")
			dtEncryptionNotBefore = g_rsClientConfiguration("EncryptionNotBefore")
			dtEncryptionNotAfter = g_rsClientConfiguration("EncryptionNotAfter")
			blnSigningSubjectMatch = g_rsClientConfiguration("SigningSubjectMatch")
			dtSigningNotBefore = g_rsClientConfiguration("SigningNotBefore")
			dtSigningNotAfter = g_rsClientConfiguration("SigningNotAfter")
			strSCCMState = g_rsClientConfiguration("CCMEXECState")
			strSCCMStartMode = g_rsClientConfiguration("CCMEXECStartMode")
			strGPClientState = g_rsClientConfiguration("GPClientState")
			strGPClientStartMode = g_rsClientConfiguration("GPClientStartMode")
			strMcFrameworkState = g_rsClientConfiguration("McFrameworkState")
			strMcFrameworkStartMode = g_rsClientConfiguration("McFrameworkStartMode")
			strMcShieldState = g_rsClientConfiguration("McShieldState")
			strMcShieldStartMode = g_rsClientConfiguration("McShieldStartMode")
			strLanmanServerState = g_rsClientConfiguration("LanmanServerState")
			strLanmanServerStartMode = g_rsClientConfiguration("LanmanServerStartMode")
			strRPCSSState = g_rsClientConfiguration("RPCSSState")
			strRPCSSStartMode = g_rsClientConfiguration("RPCSSStartMode")
			strSMSTSMGRState = g_rsClientConfiguration("SMSTaskSequenceAgentState")
			strSMSTSMGRStartMode = g_rsClientConfiguration("SMSTaskSequenceAgentStartMode")
			strRemoteRegistryState = g_rsClientConfiguration("RemoteRegistryState")
			strRemoteRegistryStartMode = g_rsClientConfiguration("RemoteRegistryStartMode")
			blnEnableDCOM = g_rsClientConfiguration("EnableDCOM")
			strManuallyAssignedSiteCode = g_rsClientConfiguration("ManuallyAssignedSiteCode")
			strEncryptionSubject = g_rsClientConfiguration("EncryptionSubject")
			intEncryptionCertNumber = g_rsClientConfiguration("EncryptionCertNumber")
			strEncryptionRegKey = g_rsClientConfiguration("EncryptionRegistryKey")
			strSigningSubject = g_rsClientConfiguration("SigningSubject")
			intSigningCertNumber = g_rsClientConfiguration("SigningCertNumber")
			strSigningRegKey = g_rsClientConfiguration("SigningRegistryKey")
			blnWMIGoodToGo = g_rsClientConfiguration("WMIGoodToGo")
			blnWMIRegGoodToGo = g_rsClientConfiguration("WMIRegGoodToGo")
			blnRemoteRegGoodToGo = g_rsClientConfiguration("RemoteRegGoodToGo")
			strOSVersion = g_rsClientConfiguration("OSVersion")
			strOSBits = g_rsClientConfiguration("AddressWidth")
			If (strOSBits = "64") Then
				blnIs64BitMachine = True
			Else
				blnIs64BitMachine = False
			End If
			'
			' Services -> Default values
			'	State = "Unknown"
			'	StartMode = "Unknown"
			'
			' Services -> Interrogation complete (Doesn't exist)
			'	State = "Not Installed"
			'	StartMode = "Not Applicable"
			'	
			' Determine if repair processing is necessary
			'
			If ((((UCase(strWUAUState) <> "UNKNOWN") And (UCase(strWUAUState) <> "NOT INSTALLED") And (UCase(strWUAUState) <> "RUNNING")) Or _
				((UCase(strWUAUStartMode) <> "AUTO") And (UCase(strWUAUStartMode) <> "NOT APPLICABLE"))) Or _
				((UCase(strBITSState) <> "UNKNOWN") And (UCase(strBITSState) <> "NOT INSTALLED") And (UCase(strBITSStartMode) <> "MANUAL") And _
					(UCase(strBITSStartMode) <> "AUTO") And (UCase(strWUAUStartMode) <> "NOT APPLICABLE")) Or _
				(((UCase(strWinMgmtState) <> "UNKNOWN") And (UCase(strWinMgmtState) <> "NOT INSTALLED") And (UCase(strWinMgmtState) <> "RUNNING")) Or _
					((UCase(strWinMgmtStartMode) <> "AUTO") And (UCase(strWinMgmtStartMode) <> "NOT APPLICABLE"))) Or _
				(((UCase(strMcFrameworkState) <> "UNKNOWN") And (UCase(strMcFrameworkState) <> "NOT INSTALLED") And (UCase(strMcFrameworkState) <> "RUNNING")) Or _
					((UCase(strMcFrameworkStartMode) <> "AUTO") And (UCase(strMcFrameworkStartMode) <> "NOT APPLICABLE"))) Or _
				(((UCase(strMcShieldState) <> "UNKNOWN") And (UCase(strMcShieldState) <> "NOT INSTALLED") And (UCase(strMcFrameworkState) <> "RUNNING")) Or _
					((UCase(strMcShieldStartMode) <> "AUTO") And (UCase(strMcShieldStartMode) <> "NOT APPLICABLE"))) Or _
				(((UCase(strGPClientState) <> "UNKNOWN") And (UCase(strGPClientState) <> "NOT INSTALLED") And (UCase(strGPClientState) <> "RUNNING")) Or _
					((UCase(strGPClientStartMode) <> "AUTO") And (UCase(strGPClientStartMode) <> "NOT APPLICABLE"))) Or _
				(((UCase(strSCCMState) <> "UNKNOWN") And (UCase(strSCCMState) <> "NOT INSTALLED") And (UCase(strSCCMState) <> "RUNNING")) Or _
					((UCase(strSCCMStartMode) <> "AUTO") And (UCase(strSCCMStartMode) <> "NOT APPLICABLE"))) Or _
				(((UCase(strLanmanServerState) <> "UNKNOWN") And (UCase(strLanmanServerState) <> "NOT INSTALLED") And (UCase(strLanmanServerState) <> "RUNNING")) Or _
					((UCase(strLanmanServerStartMode) <> "AUTO") And (UCase(strLanmanServerStartMode) <> "NOT APPLICABLE"))) Or _
				(((UCase(strRPCSSState) <> "UNKNOWN") And (UCase(strRPCSSState) <> "NOT INSTALLED") And (UCase(strRPCSSState) <> "RUNNING")) Or _
					((UCase(strRPCSSStartMode) <> "AUTO") And (UCase(strRPCSSStartMode) <> "NOT APPLICABLE"))) Or _
				((UCase(strSMSTSMGRState) <> "UNKNOWN") And (UCase(strSMSTSMGRState) <> "NOT INSTALLED") And (UCase(strSMSTSMGRStartMode) <> "MANUAL") And _
					(UCase(strSMSTSMGRStartMode) <> "NOT APPLICABLE")) Or _
				((UCase(strRemoteRegistryState) <> "UNKNOWN") And (UCase(strRemoteRegistryState) <> "NOT INSTALLED") And (UCase(strRemoteRegistryStartMode) <> "MANUAL") And _
					(UCase(strRemoteRegistryStartMode) <> "NOT APPLICABLE")) Or _
				(blnEnableDCOM = False) Or _
				((blnEncryptionSubjectMatch = False) Or (blnSigningSubjectMatch = False)) Or _
				((dtEncryptionNotBefore > Now()) Or (Now() > dtEncryptionNotAfter)) Or _
				((dtSigningNotBefore > Now()) Or (Now() > dtSigningNotAfter)) Or _
				(UCase(strManuallyAssignedSiteCode) <> UCase(g_strSiteCode))) Then
				'
				' Repair/RepairSiteCode processing is required
				'
				strRepairFile = g_strRepairOutputPath & "Repair_" & strPassedParameter & "_" & g_objFunctions.BuildDateString(Now) & ".txt"
				Set objRepairFile = g_objFSO.OpenTextFile(strRepairFile, FOR_WRITE, CREATE_IF_NON_EXISTENT)
				objRepairFile.WriteLine("Doing ClientHealth repair on " & strPassedParameter)
				If (g_blnRepair) Then
					If ((UCase(strWUAUState) <> "UNKNOWN") And (UCase(strWUAUState) <> "NOT INSTALLED") And (UCase(strWUAUStartMode) <> "NOT APPLICABLE")) Then
						'
						' Service is installed and we have data
						'
						If ((UCase(strWUAUState) <> "RUNNING") Or (UCase(strWUAUStartMode) <> "AUTO")) Then
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "WUAUSERV", "STOP", strWUAUState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "WUAUSERV", "CHANGE", strWUAUState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "WUAUSERV", "START", strWUAUState, objRepairFile)
						End If
					End If
					If ((UCase(strBITSState) <> "UNKNOWN") And (UCase(strBITSState) <> "NOT INSTALLED") And (UCase(strBITSStartMode) <> "NOT APPLICABLE")) Then
						'
						' Service is installed and we have data
						'
						If ((UCase(strBITSStartMode) <> "MANUAL") And (UCase(strBITSStartMode) <> "AUTO")) Then
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "BITS", "STOP", strBITSState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "BITS", "CHANGE", strBITSState, objRepairFile)
						End If
					End If
					If ((UCase(strWinMgmtState) <> "UNKNOWN") And (UCase(strWinMgmtState) <> "NOT INSTALLED") And (UCase(strWinMgmtStartMode) <> "NOT APPLICABLE")) Then
						'
						' Service is installed and we have data
						'
						If ((UCase(strWinMgmtState) <> "RUNNING") Or (UCase(strWinMgmtStartMode) <> "AUTO")) Then
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "WINMGMT", "STOP", strWinMgmtState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "WINMGMT", "CHANGE", strWinMgmtState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "WINMGMT", "START", strWinMgmtState, objRepairFile)
						End If
					End If
					If ((UCase(strGPClientState) <> "UNKNOWN") And (UCase(strGPClientState) <> "NOT INSTALLED") And (UCase(strGPClientStartMode) <> "NOT APPLICABLE")) Then
						'
						' Service is installed and we have data
						'
						If ((UCase(strGPClientState) <> "RUNNING") Or (UCase(strGPClientStartMode) <> "AUTO")) Then
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "GPSVC", "STOP", strGPClientState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "GPSVC", "CHANGE", strGPClientState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "GPSVC", "START", strGPClientState, objRepairFile)
						End If
					End If
					If ((UCase(strMcFrameworkState) <> "UNKNOWN") And (UCase(strMcFrameworkState) <> "NOT INSTALLED") And (UCase(strMcFrameworkStartMode) <> "NOT APPLICABLE")) Then
						'
						' Service is installed and we have data
						'
						If ((UCase(strMcFrameworkState) <> "RUNNING") Or (UCase(strMcFrameworkStartMode) <> "AUTO")) Then
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "MCAFEEFRAMEWORK", "STOP", strMcFrameworkState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "MCAFEEFRAMEWORK", "CHANGE", strMcFrameworkState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "MCAFEEFRAMEWORK", "START", strMcFrameworkState, objRepairFile)
						End If
					End If
					If ((UCase(strMcShieldState) <> "UNKNOWN") And (UCase(strMcShieldState) <> "NOT INSTALLED") And (UCase(strMcShieldStartMode) <> "NOT APPLICABLE")) Then
						'
						' Service is installed and we have data
						'
						If ((UCase(strMcShieldState) <> "RUNNING") Or (UCase(strMcShieldStartMode) <> "AUTO")) Then
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "MCSHIELD", "STOP", strMcShieldState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "MCSHIELD", "CHANGE", strMcShieldState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "MCSHIELD", "START", strMcShieldState, objRepairFile)
						End If
					End If
					If ((UCase(strLanmanServerState) <> "UNKNOWN") And (UCase(strLanmanServerState) <> "NOT INSTALLED") And (UCase(strLanmanServerStartMode) <> "NOT APPLICABLE")) Then
						'
						' Service is installed and we have data
						'
						If ((UCase(strLanmanServerState) <> "RUNNING") Or (UCase(strLanmanServerStartMode) <> "AUTO")) Then
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "LANMANSERVER", "STOP", strLanmanServerState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "LANMANSERVER", "CHANGE", strLanmanServerState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "LANMANSERVER", "START", strLanmanServerState, objRepairFile)
						End If
					End If
					If ((UCase(strRPCSSState) <> "UNKNOWN") And (UCase(strRPCSSState) <> "NOT INSTALLED") And (UCase(strRPCSSStartMode) <> "NOT APPLICABLE")) Then
						'
						' Service is installed and we have data
						'
						If ((UCase(strRPCSSState) <> "RUNNING") Or (UCase(strRPCSSStartMode) <> "AUTO")) Then
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "RPCSS", "STOP", strRPCSSState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "RPCSS", "CHANGE", strRPCSSState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "RPCSS", "START", strRPCSSState, objRepairFile)
						End If
					End If
					If ((UCase(strSMSTSMGRState) <> "UNKNOWN") And (UCase(strSMSTSMGRState) <> "NOT INSTALLED") And (UCase(strSMSTSMGRStartMode) <> "NOT APPLICABLE")) Then
						'
						' Service is installed and we have data
						'
						If (UCase(strSMSTSMGRStartMode) <> "MANUAL") Then
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "SMSTSMGR", "STOP", strSMSTSMGRState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "SMSTSMGR", "CHANGE", strSMSTSMGRState, objRepairFile)
						End If
					End If
					If ((UCase(strRemoteRegistryState) <> "UNKNOWN") And (UCase(strRemoteRegistryState) <> "NOT INSTALLED") And (UCase(strRemoteRegistryStartMode) <> "NOT APPLICABLE")) Then
						'
						' Service is installed and we have data
						'
						If ((UCase(strRemoteRegistryState) <> "RUNNING") Or (UCase(strRemoteRegistryStartMode) <> "AUTO")) Then
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "REMOTEREGISTRY", "STOP", strRemoteRegistryState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "REMOTEREGISTRY", "CHANGE", strRemoteRegistryState, objRepairFile)
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "REMOTEREGISTRY", "START", strRemoteRegistryState, objRepairFile)
						End If
					End If
					'
					' Process CCMExec/Certificates only if CCMExec is installed
					'
					blnRestartRequired = False
					If ((UCase(strSCCMState) <> "UNKNOWN") And (UCase(strSCCMState) <> "NOT INSTALLED") And (UCase(strSCCMStartMode) <> "NOT APPLICABLE")) Then
						'
						' Service is installed and we have data
						'
						If (((blnEncryptionSubjectMatch = False) And ((intEncryptionCertNumber <> -1) Or (strEncryptionRegKey <> "FFFFFFFF"))) Or _
							((blnSigningSubjectMatch = False) And ((intSigningCertNumber <> -1) Or (strSigningRegKey <> "FFFFFFFF"))) Or _
							((dtEncryptionNotBefore > Now()) Or (Now() > dtEncryptionNotAfter)) Or _
							((dtSigningNotBefore > Now()) Or (Now() > dtSigningNotAfter)) Or _
							((UCase(strSCCMState) = "RUNNING") And (UCase(strSCCMStartMode) <> "AUTO"))) Then
							'
							' Stop the service
							'
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "CCMEXEC", "STOP", strSCCMState, objRepairFile)
						End If
						If ((UCase(strSCCMStartMode) <> "AUTO") And (UCase(strSCCMStartMode) <> "NOT APPLICABLE")) Then
							Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "CCMEXEC", "CHANGE", strSCCMState, objRepairFile)
						End If
						blnRestartRequired = True
					End If
					'
					' CCMExec service is stopped or not installed - update the certs regardless
					'
					If (((blnEncryptionSubjectMatch = False) And ((intEncryptionCertNumber <> -1) Or (strEncryptionRegKey <> "FFFFFFFF"))) Or _
						((dtEncryptionNotBefore > Now()) Or (Now() > dtEncryptionNotAfter))) Then
						'
						' Delete the Encryption certificate (if it exists)
						'
						Call CorrectCertificate(objRemoteRegServer, blnWMIRegGoodToGo, blnRemoteRegGoodToGo, strWMIConnectedWithThis, _
													blnIs64BitMachine, strOSVersion, "Encryption", intEncryptionCertNumber, _
													strEncryptionRegKey, objRepairFile)
					End If
					If (((blnSigningSubjectMatch = False) And ((intSigningCertNumber <> -1) Or (strSigningRegKey <> "FFFFFFFF"))) Or _
						((dtSigningNotBefore > Now()) Or (Now() > dtSigningNotAfter))) Then
						'
						' Delete the Sigining certificate (if it exists)
						'
						Call CorrectCertificate(objRemoteRegServer, blnWMIRegGoodToGo, blnRemoteRegGoodToGo, strWMIConnectedWithThis, _
													blnIs64BitMachine, strOSVersion, "Signing", intSigningCertNumber, _
													strSigningRegKey, objRepairFile)
					End If
					If (blnRestartRequired) Then
						Call UpdateService(objRemoteWMIServer, blnWMIGoodToGo, strWMIConnectedWithThis, "CCMEXEC", "START", strSCCMState, objRepairFile)
					End If
				End If
				'
				' Enable DCOM
				'
				If (blnEnableDCOM = False) Then
					objRepairFile.WriteLine("Setting registry value for EnableDCOM to True")
					strRegistryHive = "HKLM"
					strRegistryKey = "SOFTWARE\Microsoft\OLE"
					strEntryName = "EnableDCOM"
					strOperation = "String"
					If (strOSBits = "64") Then
						blnIs64BitMachine = True
					Else
						blnIs64BitMachine = False
					End If
					blnIsWow6432Node = True
					intRetVal = g_objRegistryProcessing.SetRegistryEntry(objRemoteRegServer, strWMIConnectedWithThis, blnWMIRegGoodToGo, _
																			blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strEntryName, _
																			"Y", strOperation, blnIs64BitMachine, blnIsWow6432Node, g_objLogAndTrace, _
																			g_objLogAndTraceErrors)
					If (intRetVal = 0) Then
						objRepairFile.WriteLine("Registry value update complete - reboot required for setting to take effect")
					Else
						objRepairFile.WriteLine("Registry value update failed")
					End If
				End If
				If (g_blnRepairSiteCode) Then
					'
					' Update site code
					'
					If (UCase(strManuallyAssignedSiteCode) <> UCase(g_strSiteCode)) Then
						objRepairFile.WriteLine("Setting registry value for AssignedSiteCode to " & g_strSiteCode)
						strRegistryHive = "HKLM"
						strRegistryKey = "SOFTWARE\Microsoft\SMS\Mobile Client"
						strEntryName = "AssignedSiteCode"
						strOperation = "String"
						If (strOSBits = "64") Then
							blnIs64BitMachine = True
						Else
							blnIs64BitMachine = False
						End If
						blnIsWow6432Node = True
						intRetVal = g_objRegistryProcessing.SetRegistryEntry(objRemoteRegServer, strWMIConnectedWithThis, blnWMIRegGoodToGo, _
																				blnRemoteRegGoodToGo, strRegistryHive, strRegistryKey, strEntryName, _
																				g_strSiteCode, strOperation, blnIs64BitMachine, blnIsWow6432Node, _
																				g_objLogAndTrace, g_objLogAndTraceErrors)
						If (intRetVal = 0) Then
							objRepairFile.WriteLine("Registry value update complete")
						Else
							objRepairFile.WriteLine("Registry value update failed")
						End If
					End If
				End If
				objRepairFile.WriteLine("Updates completed on " & strPassedParameter)
				objRepairFile.Close
				Set objRepairFile = Nothing
				'
				' Rescan the computer to get up-to-date results
				'
				g_objFunctions.DeleteAllRecordsetRows(g_rsConnectionStatus)
				g_objFunctions.DeleteAllRecordsetRows(g_rsClientConfiguration)
				Call ScanClientConfiguration(strPassedParameter, intFlag)
			Else
				'
				' Repair processing is not necessary - write the updated ClientHealth information to the XML file.
				'
				Call BuildClientHealthXML(strPassedParameter, dtStartTime)
			End If
		End If
	End If

End Function

Function DisplayHelpMessage()
'*****************************************************************************************************************************************
'*  Purpose:				Displays the command line help
'*  Called by:				Main
'*  Comments:				-
'*****************************************************************************************************************************************
	WScript.Echo ""
	WScript.Echo "SYNTAX: CScript " & WScript.ScriptName & " [/Switch] [/Switch]..."
	WScript.Echo ""
	WScript.Echo "  Valid Switch values:"
	WScript.Echo ""
	WScript.Echo "  /Machine"
	WScript.Echo "     FQDN, NetBIOS name, or IP Address of machine to be processed"
	WScript.Echo ""
	WScript.Echo "  /File"
	WScript.Echo "     Input file containing FQDNs, NetBIOS Names, or IP Addresses to be processed"
	WScript.Echo ""
	WScript.Echo "  /Repair"
	WScript.Echo "     Execute the repair process against all machines passed"
	WScript.Echo ""
	WScript.Echo "  /RepairSiteCode"
	WScript.Echo "     Execute the repair process against all machines passed (SiteCode)"
	WScript.Echo ""
	WScript.Echo "  /Scan"
	WScript.Echo "     Gather (scan) the SCCM configuration from all machines passed"
	WScript.Echo ""
	WScript.Echo "  /ProcessXML_SCCM"
	WScript.Echo "     Process all XML files and load SCCM data into tab delimited file"
	WScript.Echo ""
	WScript.Echo "  /ProcessXML_HBSS"
	WScript.Echo "     Process all XML files and load HBSS data into tab delimited file"
	WScript.Echo ""
	WScript.Echo "  /SiteCode"
	WScript.Echo "     Used with /RepairSiteCode processing to update registry with correct value"
	WScript.Echo ""
	WScript.Echo "  /Verbose"
	WScript.Echo "     Used with /ProcessXML_SCCM processing to display ALL collected information"
	WScript.Echo "         The output file was getting too cluttered so it was pared down."
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
	WScript.Echo "  CScript " & WScript.ScriptName & " /Machine:SERVER2003EEPDC.test.local /Scan"
	WScript.Echo "  CScript " & WScript.ScriptName & " /File:D:\Input.txt /Scan"
	WScript.Echo "  CScript " & WScript.ScriptName & " /ProcessXML_SCCM"
	WScript.Echo "  CScript " & WScript.ScriptName & " /File:D:\Input.txt /Repair /SiteCode:SCU"
	WScript.Echo "  CScript " & WScript.ScriptName & " /?"
	WScript.Echo ""
	WScript.Echo "NOTE: If a parameter contains a space it MUST be within quotes"
	WScript.Echo "  CScript " & WScript.ScriptName & " /File:" & Chr(34) & "D:\Temp Path\xyz.txt" & Chr(34)
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

	strWSHVersion = objFSO.GetFileVersion(WScript.FullName)
	If (CStr(strWSHVersion) < CStr(strWSHVersionAtLeast)) Then
		MsgBox("The version of Windows Scripting Host (WSH) is not correct.  Download a version newer than " & strWSHVersionAtLeast & VbCrLf & _
				" from Microsoft at http://www.microsoft.com/downloads/results.aspx?displaylang=en&freeText=windows+script+host.")
		WScript.Sleep 1000
		WScript.Quit
	End If
	'
	' Force the script to run using CScript
	'
	If (Right(UCase(WScript.FullName), 11) = "WSCRIPT.EXE") Then
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
Set g_objXMLProcessing = New XMLProcessing
Set g_objPossibleProcessing = New PossibleProcessing
Set g_objClientResolution = New ClientResolution
Set g_objRegistryProcessing = New RegistryProcessing
'
' Create XML variables
'
Set g_xmlDoc = CreateObject("Microsoft.XMLDOM")
'
' Create the PassedParams and ProcessingInfo child elements
'
Call g_objXMLProcessing.CreateNewElement(g_xmlDoc, g_xmlElementPassedParams, "PassedParams")
'
' Set default values
'
g_strMachine = ""
g_blnPassedMachine = False
g_strSourceFile = ""
g_blnPassedFile = False
g_blnRepair = False
g_blnRepairSiteCode = False
g_blnScan = False
g_blnProcessXML_SCCM = False
g_blnProcessXML_HBSS = False
g_strSiteCode = ""
g_blnVerbose = False
g_blnTraceAll = False
g_blnTraceBasic = False
g_blnTracePossibleProcessing = False
g_blnTraceRegistryProcessing = False
g_blnTraceClientResolution = False
g_blnTraceLoadRS = False
g_blnTraceExecCmdGeneric = False
g_blnTraceXMLFileAndRegistry = False
g_strArgument = ""
g_dtCreationTimestamp = g_objFunctions.GetGMTTimestamp()
'
' Process command line arguments
'
For Each g_strArgument In Wscript.Arguments.Named
	'
	' Process the command line parameters
	'
	Select Case UCase(g_strArgument)
		Case "MACHINE"
			g_strMachine = WScript.Arguments.Named(g_strArgument)
			g_blnPassedMachine = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "Machine", g_strMachine)
		Case "FILE"
			g_strSourceFile = WScript.Arguments.Named(g_strArgument)
			g_blnPassedFile = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "File", g_strSourceFile)
		Case "REPAIR"
			g_blnRepair = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "Repair", g_blnRepair)
		Case "REPAIRSITECODE"
			g_blnRepairSiteCode = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "RepairSiteCode", g_blnRepairSiteCode)
		Case "SCAN"
			g_blnScan = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "Scan", g_blnScan)
		Case "PROCESSXML_SCCM"
			g_blnProcessXML_SCCM = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "ProcessXML_SCCM", g_blnProcessXML_SCCM)
		Case "PROCESSXML_HBSS"
			g_blnProcessXML_HBSS = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "ProcessXML_HBSS", g_blnProcessXML_HBSS)
		Case "SITECODE"
			g_strSiteCode = WScript.Arguments.Named(g_strArgument)
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "SiteCode", g_strSiteCode)
		Case "VERBOSE"
			g_blnVerbose = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "Verbose", g_blnVerbose)
		Case "TRACEALL"
			g_blnTraceBasic = True
			g_blnTracePossibleProcessing = True
			g_blnTraceRegistryProcessing = True
			g_blnTraceClientResolution = True
			g_blnTraceLoadRS = True
			g_blnTraceExecCmdGeneric = True
			g_blnTraceXMLFileAndRegistry = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "TraceAll", g_blnTraceAll)
		Case "TRACEBASIC"
			g_blnTraceBasic = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "TraceBasic", g_blnTraceBasic)
		Case "TRACEPOSSIBLEPROCESSING"
			g_blnTracePossibleProcessing = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "TracePossibleProcessing", g_blnTracePossibleProcessing)
		Case "TRACEREGISTRYPROCESSING"
			g_blnTraceRegistryProcessing = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "TraceRegistryProcessing", g_blnTraceRegistryProcessing)
		Case "TRACECLIENTRESOLUTION"
			g_blnTraceClientResolution = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "TraceClientResolution", g_blnTraceClientResolution)
		Case "TRACELOADRS"
			g_blnTraceLoadRS = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "TraceLoadRS", g_blnTraceLoadRS)
		Case "TRACEEXECCMDGENERIC"
			g_blnTraceExecCmdGeneric = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "TraceExecCmdGeneric", g_blnTraceExecCmdGeneric)
		Case "TRACEXML"
			g_blnTraceXMLFileAndRegistry = True
			Call g_objXMLProcessing.CreateSetAndLinkAttribute(g_xmlDoc, g_xmlElementPassedParams, "TraceXMLFileAndRegistry", g_blnTraceXMLFileAndRegistry)
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
' Check for Machine or File parameter
'
Set g_objShell = CreateObject("WScript.Shell")
g_blnProcessingLocal = False
If ((g_blnRepair) Or (g_blnScan) Or (g_blnRepairSiteCode)) Then
	g_strThisComputer = g_objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
	If ((g_strMachine = "") And (g_strSourceFile = "")) Then
		WScript.Echo "Machine and File not passed...defaulting to local machine..."
		g_strMachine = g_strThisComputer
		g_blnProcessingLocal = True
	End If
	If (InStr(1, g_strMachine, g_strThisComputer, vbTextCompare) > 0) Then
		g_blnProcessingLocal = True
	End If
End If
Set g_objShell = Nothing
'
' Ensure there is processing to do
'
If ((g_blnRepair = False) And (g_blnScan = False) And (g_blnProcessXML_SCCM = False) And (g_blnProcessXML_HBSS = False) And _
	(g_blnRepairSiteCode = False)) Then
	WScript.Echo "A processing request (Repair, RepairSiteCode, Validate, and/or ProcessXML) must be made.  Please try again."
	WScript.Quit
End If
'
' Ensure SiteCode was passed for Repair processing
'
If ((g_blnRepairSiteCode) And (g_strSiteCode = "")) Then
	WScript.Echo "A valid site code is required for RepairSiteCode processing.  Please try again."
	WScript.Quit
End If
g_strParentFolder = g_objFunctions.GetParentFolder()
'
' Create the MachinesToProcess recordset
'
Set g_rsMachinesToProcess = CreateObject("ADODB.Recordset")
g_rsMachinesToProcess.Fields.Append "Machine", adVarWChar, 80
g_rsMachinesToProcess.Open
'
' Create the Generic recordset needed for ExecCmdGeneric processing
'
Set g_rsGeneric = CreateObject("ADODB.Recordset")
g_rsGeneric.Fields.Append "SavedData", adVarChar, 255
g_rsGeneric.Open
'
' Create the ClientConfiguration recordset
'
Set g_rsClientConfiguration = CreateObject("ADODB.Recordset")
g_rsClientConfiguration.Fields.Append "ComputerFQDN", adVarChar, 80
g_rsClientConfiguration.Fields.Append "ComputerName", adVarChar, 30
g_rsClientConfiguration.Fields.Append "IPAddress", adVarChar, 15
g_rsClientConfiguration.Fields.Append "OSName", adVarChar, 100
g_rsClientConfiguration.Fields.Append "OSBuildNumber", adVarChar, 50
g_rsClientConfiguration.Fields.Append "LastBootupTime", adDate
g_rsClientConfiguration.Fields.Append "OSBuildType", adVarChar, 50
g_rsClientConfiguration.Fields.Append "OSType", adInteger
g_rsClientConfiguration.Fields.Append "OSTypeText", adVarChar, 50
g_rsClientConfiguration.Fields.Append "ProductType", adInteger
g_rsClientConfiguration.Fields.Append "ProductTypeText", adVarChar, 50
g_rsClientConfiguration.Fields.Append "OSVersion", adVarChar, 20
g_rsClientConfiguration.Fields.Append "AddressWidth", adSmallInt
g_rsClientConfiguration.Fields.Append "PercentFreeC", adDouble
g_rsClientConfiguration.Fields.Append "ManuallyAssignedSiteCode", adVarChar, 10
g_rsClientConfiguration.Fields.Append "GPOAssignedSiteCode", adVarChar, 10
g_rsClientConfiguration.Fields.Append "EncryptionSubject", adVarChar, 255
g_rsClientConfiguration.Fields.Append "EncryptionSubjectMatch", adBoolean
g_rsClientConfiguration.Fields.Append "EncryptionCertNumber", adSmallInt
g_rsClientConfiguration.Fields.Append "EncryptionRegistryKey", adVarChar, 255
g_rsClientConfiguration.Fields.Append "EncryptionNotBefore", adDate
g_rsClientConfiguration.Fields.Append "EncryptionNotAfter", adDate
g_rsClientConfiguration.Fields.Append "SigningSubject", adVarChar, 255
g_rsClientConfiguration.Fields.Append "SigningSubjectMatch", adBoolean
g_rsClientConfiguration.Fields.Append "SigningCertNumber", adSmallInt
g_rsClientConfiguration.Fields.Append "SigningRegistryKey", adVarChar, 255
g_rsClientConfiguration.Fields.Append "SigningNotBefore", adDate
g_rsClientConfiguration.Fields.Append "SigningNotAfter", adDate
g_rsClientConfiguration.Fields.Append "WUAUState", adVarChar, 20
g_rsClientConfiguration.Fields.Append "WUAUStartMode", adVarChar, 20
g_rsClientConfiguration.Fields.Append "CCMEXECState", adVarChar, 20
g_rsClientConfiguration.Fields.Append "CCMEXECStartMode", adVarChar, 20
g_rsClientConfiguration.Fields.Append "BITSState", adVarChar, 20
g_rsClientConfiguration.Fields.Append "BITSStartMode", adVarChar, 20
g_rsClientConfiguration.Fields.Append "WinMgmtState", adVarChar, 20
g_rsClientConfiguration.Fields.Append "WinMgmtStartMode", adVarChar, 20
g_rsClientConfiguration.Fields.Append "GPClientState", adVarChar, 20
g_rsClientConfiguration.Fields.Append "GPClientStartMode", adVarChar, 20
g_rsClientConfiguration.Fields.Append "CCMSetupState", adVarChar, 20
g_rsClientConfiguration.Fields.Append "CCMSetupStartMode", adVarChar, 20
g_rsClientConfiguration.Fields.Append "LanmanServerState", adVarChar, 20
g_rsClientConfiguration.Fields.Append "LanmanServerStartMode", adVarChar, 20
g_rsClientConfiguration.Fields.Append "RPCSSState", adVarChar, 20
g_rsClientConfiguration.Fields.Append "RPCSSStartMode", adVarChar, 20
g_rsClientConfiguration.Fields.Append "SMSTaskSequenceAgentState", adVarChar, 20
g_rsClientConfiguration.Fields.Append "SMSTaskSequenceAgentStartMode", adVarChar, 20
g_rsClientConfiguration.Fields.Append "RemoteRegistryState", adVarChar, 20
g_rsClientConfiguration.Fields.Append "RemoteRegistryStartMode", adVarChar, 20
g_rsClientConfiguration.Fields.Append "SCCMVersion", adVarChar, 20
g_rsClientConfiguration.Fields.Append "WMIGoodToGo", adBoolean
g_rsClientConfiguration.Fields.Append "WMIRegGoodToGo", adBoolean
g_rsClientConfiguration.Fields.Append "RemoteRegGoodToGo", adBoolean
g_rsClientConfiguration.Fields.Append "WindowsUpdateServer", adVarChar, 150
g_rsClientConfiguration.Fields.Append "NameServers", adVarChar, 150
g_rsClientConfiguration.Fields.Append "LastLoggedOnUser", adVarChar, 75
g_rsClientConfiguration.Fields.Append "PendingReboot", adBoolean
g_rsClientConfiguration.Fields.Append "SMSUID", adChar, 36
g_rsClientConfiguration.Fields.Append "PreviousSMSUID", adChar, 36
g_rsClientConfiguration.Fields.Append "LastChangedSMSUID", adDate
g_rsClientConfiguration.Fields.Append "ClientSite", adVarChar, 20
g_rsClientConfiguration.Fields.Append "CurrentMP", adVarChar, 80
g_rsClientConfiguration.Fields.Append "ADSiteName", adVarChar, 20
g_rsClientConfiguration.Fields.Append "SDCVersion", adVarChar, 30
g_rsClientConfiguration.Fields.Append "WUAVersion", adVarChar, 75
g_rsClientConfiguration.Fields.Append "RunningAdvertisements", adInteger
g_rsClientConfiguration.Fields.Append "MostRecentSMSFolder", adVarChar, 200
g_rsClientConfiguration.Fields.Append "MostRecentSMSFolderDate", adDate
g_rsClientConfiguration.Fields.Append "EnableDCOM", adBoolean
g_rsClientConfiguration.Fields.Append "SCCMHealthy", adBoolean
'
' HBSS settings
'
g_rsClientConfiguration.Fields.Append "AVDatDate", adDate
g_rsClientConfiguration.Fields.Append "CatalogVersionDate", adDate
g_rsClientConfiguration.Fields.Append "SiteListName", adVarChar, 255
g_rsClientConfiguration.Fields.Append "SiteListIP", adVarChar, 255
g_rsClientConfiguration.Fields.Append "SiteListPort", adVarChar, 100
g_rsClientConfiguration.Fields.Append "EPORegistryName", adVarChar, 255
g_rsClientConfiguration.Fields.Append "EPORegistryIP", adVarChar, 255
g_rsClientConfiguration.Fields.Append "EPORegistryPort", adVarChar, 100
g_rsClientConfiguration.Fields.Append "AgentGUID", adVarChar, 100
g_rsClientConfiguration.Fields.Append "McFrameworkState", adVarChar, 20
g_rsClientConfiguration.Fields.Append "McFrameworkStartMode", adVarChar, 20
g_rsClientConfiguration.Fields.Append "McShieldState", adVarChar, 20
g_rsClientConfiguration.Fields.Append "McShieldStartMode", adVarChar, 20
g_rsClientConfiguration.Fields.Append "LastASCTime", adDate
g_rsClientConfiguration.Fields.Append "PropsVersionDate", adDate
g_rsClientConfiguration.Fields.Append "AgentWakeUpPort", adInteger
g_rsClientConfiguration.Fields.Append "MAVersion", adVarChar, 25
g_rsClientConfiguration.Fields.Append "FWEnabled", adBoolean
g_rsClientConfiguration.Fields.Append "VSEVersion", adVarChar, 25
g_rsClientConfiguration.Fields.Append "HIPSVersion", adVarChar, 25
g_rsClientConfiguration.Fields.Append "DLPVersion", adVarChar, 25
g_rsClientConfiguration.Fields.Append "HBSSHealthy", adBoolean
g_rsClientConfiguration.Open
'
' Create the ConnectionStatus structure
'
Set g_rsConnectionStatus = CreateObject("ADODB.Recordset")
g_rsConnectionStatus.Fields.Append "PassedParameter", adVarChar, 80
g_rsConnectionStatus.Fields.Append "HostAlive", adBoolean
g_rsConnectionStatus.Fields.Append "WMIDCOMConnectionSuccessful", adBoolean
g_rsConnectionStatus.Fields.Append "WMIConnectedWithThis", adVarChar, 80
g_rsConnectionStatus.Fields.Append "WMIConnectErrorOccurred", adBoolean
g_rsConnectionStatus.Fields.Append "WMIConnectError", adLongVarWChar, 512
g_rsConnectionStatus.Fields.Append "WMIProcessingGoodToGo", adBoolean
g_rsConnectionStatus.Fields.Append "RegistryProcessingMethodUsed", adVarChar, 25
g_rsConnectionStatus.Fields.Append "WMIRegistryDCOMConnectionSuccessful", adBoolean
g_rsConnectionStatus.Fields.Append "WMIRegistryConnectedWithThis", adVarChar, 80
g_rsConnectionStatus.Fields.Append "WMIRegistryConnectErrorOccurred", adBoolean
g_rsConnectionStatus.Fields.Append "WMIRegistryConnectError", adLongVarWChar, 512
g_rsConnectionStatus.Fields.Append "WMIRegistryProcessingGoodToGo", adBoolean
g_rsConnectionStatus.Fields.Append "RemoteRegistryConnectedWithThis", adVarChar, 80
g_rsConnectionStatus.Fields.Append "RemoteRegistryProcessingGoodToGo", adBoolean
g_rsConnectionStatus.Fields.Append "LocalAccountsConnectedWithThis", adVarChar, 80
g_rsConnectionStatus.Fields.Append "LocalAccountProcessingGoodToGo", adBoolean
g_rsConnectionStatus.Fields.Append "ProcessingLocal", adBoolean
g_rsConnectionStatus.Fields.Append "IPAddressAvailabilityMethod", adVarChar, 50
g_rsConnectionStatus.Fields.Append "PassedDNSHostName", adBoolean
g_rsConnectionStatus.Fields.Append "PassedHostName", adBoolean
g_rsConnectionStatus.Fields.Append "PassedIPAddress", adBoolean
g_rsConnectionStatus.Fields.Append "ResolvedDNSHostName", adVarChar, 80
g_rsConnectionStatus.Fields.Append "ResolvedHostName", adVarChar, 30
g_rsConnectionStatus.Fields.Append "ResolvedNetBIOSName", adVarChar, 15
g_rsConnectionStatus.Fields.Append "ResolvedIPAddress", adVarChar, 15
g_rsConnectionStatus.Fields.Append "DNSHostNameResolved", adBoolean
g_rsConnectionStatus.Fields.Append "HostNameResolved", adBoolean
g_rsConnectionStatus.Fields.Append "IPAddressResolved", adBoolean
g_rsConnectionStatus.Fields.Append "NetBIOSNameResolved", adBoolean
g_rsConnectionStatus.Fields.Append "AbortedProcessingText", adVarChar, 150
g_rsConnectionStatus.Fields.Append "ComputerFQDN", adVarChar, 80
g_rsConnectionStatus.Fields.Append "IsPartialCollection", adBoolean
g_rsConnectionStatus.Open
'
' Create the Certificates structure
'
Set g_rsCertificates = CreateObject("ADODB.Recordset")
g_rsCertificates.Fields.Append "SHA1Hash", adVarChar, 50
g_rsCertificates.Fields.Append "SerialNumber", adVarChar, 50
g_rsCertificates.Fields.Append "Issuer", adVarChar, 150
g_rsCertificates.Fields.Append "NotBefore", adDate
g_rsCertificates.Fields.Append "NotAfter", adDate
g_rsCertificates.Fields.Append "Subject", adVarChar, 255
g_rsCertificates.Fields.Append "CertNumber", adInteger
g_rsCertificates.Fields.Append "RegistryKey", adVarChar, 150
g_rsCertificates.Fields.Append "Type", adVarChar, 20
g_rsCertificates.Open
'
' Add an individual machine if it was passed.
'
If ((g_blnScan) Or (g_blnRepair) Or (g_blnRepairSiteCode)) Then
	If (g_strMachine <> "") Then
		g_rsMachinesToProcess.AddNew
		g_rsMachinesToProcess("Machine") = g_strMachine
		g_rsMachinesToProcess.Update
	End If
	If (g_strSourceFile <> "") Then
		If (g_objFSO.FileExists(g_strSourceFile)) Then
			Set g_objSourceFile = g_objFSO.OpenTextFile(g_strSourceFile)
			While Not g_objSourceFile.AtEndOfStream
				g_strMachine = UCase(Trim(g_objSourceFile.ReadLine))
				If (g_strMachine <> "") Then
					g_rsMachinesToProcess.AddNew
					g_rsMachinesToProcess("Machine") = g_strMachine
					g_rsMachinesToProcess.Update
				End If
			Wend
		End If
	End If
	'
	' Make sure there are machines to process
	'
	If (g_rsMachinesToProcess.RecordCount = 0) Then
		WScript.Echo "No Machines specified for processing.  Please try again."
		WScript.Quit
	End If
End If
g_intFlag = wbemFlagReturnWhenComplete
'
' Setup folder names
'
g_strXMLOutputPath = g_strParentFolder & "XMLOutput_ClientHealth\"
g_strErrorOutputPath = g_strParentFolder & "ErrorOutput_ClientHealth\"
g_strTraceOutputPath = g_strParentFolder & "TraceOutput_ClientHealth\"
'
' What processing was selected?
'
If ((g_blnScan) Or (g_blnRepair) Or (g_blnRepairSiteCode)) Then
	'
	' Create the "XMLOutput" folder
	'
	Call g_objFunctions.BuildPath(g_strXMLOutputPath)
	'
	' Create the "ErrorOutput" folder
	'
	Call g_objFunctions.BuildPath(g_strErrorOutputPath)
	'
	' Create the "TraceOutput" folder
	'
	Call g_objFunctions.BuildPath(g_strTraceOutputPath)
End If

If (g_blnScan) Then
	'
	' Process the records
	'
	g_rsMachinesToProcess.MoveFirst
	While Not g_rsMachinesToProcess.EOF
		g_strMachine = g_rsMachinesToProcess("Machine")
		'
		' Setup the name of the logging file
		'
		g_strLoggingFile = g_strMachine
		'
		' Setup Logging and Tracing files
		'
		g_strGUID = g_objFunctions.CreateGloballyUniqueID()
		g_blnCreateNewFile = True
		g_strLogAndTrace = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_FILE
		g_strLogAndTraceErrors = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_ERROR_FILE
		g_strLogAndTraceClientResolution = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_CLIENT_RESOLUTION_FILE
		g_strLogAndTracePossibleProcessing = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_POSSIBLE_PROCESSING_FILE
		g_strLogAndTraceRegistryProcessing = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_REGISTRY_PROCESSING_FILE
		g_strLogAndTraceExecCmdGeneric = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_EXEC_CMD_GENERIC_FILE
		g_strLogAndTraceLoadRS = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_LOADRS_FILE
		g_strLogAndTraceXMLFileAndRegistry = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_XML_FILE_AND_REGISTRY
		'
		' Setup global access to local function logging classes
		'
		Set g_objLogAndTrace = (New LoggingAndTracing)(g_strLogAndTrace, g_blnTraceBasic, g_blnCreateNewFile)
		Set g_objLogAndTraceErrors = (New LoggingAndTracing)(g_strLogAndTraceErrors, True, g_blnCreateNewFile)
		Set g_objLogAndTraceClientResolution = (New LoggingAndTracing)(g_strLogAndTraceClientResolution, g_blnTraceClientResolution, g_blnCreateNewFile)
		Set g_objLogAndTracePossibleProcessing = (New LoggingAndTracing)(g_strLogAndTracePossibleProcessing, g_blnTracePossibleProcessing, g_blnCreateNewFile)
		Set g_objLogAndTraceRegistryProcessing = (New LoggingAndTracing)(g_strLogAndTraceRegistryProcessing, g_blnTraceRegistryProcessing, g_blnCreateNewFile)
		Set g_objLogAndTraceExecCmdGeneric = (New LoggingAndTracing)(g_strLogAndTraceExecCmdGeneric, g_blnTraceExecCmdGeneric, g_blnCreateNewFile)
		Set g_objLogAndTraceLoadRS = (New LoggingAndTracing)(g_strLogAndTraceLoadRS, g_blnTraceLoadRS, g_blnCreateNewFile)
		Set g_objLogAndTraceXMLFileAndRegistry = (New LoggingAndTracing)(g_strLogAndTraceXMLFileAndRegistry, g_blnTraceXMLFileAndRegistry, g_blnCreateNewFile)
		'
		' Do the processing
		'		
		Call ScanClientConfiguration(g_strMachine, g_intFlag)
		g_objFunctions.DeleteAllRecordsetRows(g_rsGeneric)
		g_objFunctions.DeleteAllRecordsetRows(g_rsConnectionStatus)
		g_objFunctions.DeleteAllRecordsetRows(g_rsClientConfiguration)
		'
		' Clean up logging and tracing in preparation for next machine
		'
		Set g_objLogAndTrace = Nothing
		Set g_objLogAndTraceErrors = Nothing
		Set g_objLogAndTraceClientResolution = Nothing
		Set g_objLogAndTracePossibleProcessing = Nothing
		Set g_objLogAndTraceRegistryProcessing = Nothing
		Set g_objLogAndTraceExecCmdGeneric = Nothing
		Set g_objLogAndTraceLoadRS = Nothing
		Set g_objLogAndTraceXMLFileAndRegistry = Nothing
		Set g_xmlElementConnectionStatus = Nothing
		'
		' Process next machine
		'
		g_rsMachinesToProcess.MoveNext
	Wend
End If

If ((g_blnProcessXML_SCCM) Or (g_blnProcessXML_HBSS)) Then
	'
	' Ensure XMLOutput folder exists
	'
	If (Not g_objFSO.FolderExists(g_strXMLOutputPath)) Then
		WScript.Echo "XMLOutput folder doesn't exist.  Please perform a Scan prior to selecting ProcessXML."
		WScript.Quit
	End If
	'
	' Ensure there are files to process
	'
	Set g_objFolder = g_objFSO.GetFolder(g_strXMLOutputPath)
	Set g_colFiles = g_objFolder.Files
	If (g_colFiles.Count = 0) Then
		WScript.Echo "XMLOutput folder contains no XML files to process."
		WScript.Echo "Please move .XML files created during Scan processing to XMLOutput folder and try again."
		WScript.Quit
	End If
	Call ProcessXMLFiles(g_colFiles, g_blnProcessXML_SCCM, g_blnProcessXML_HBSS, g_blnVerbose)
End If
If ((g_blnRepair) Or (g_blnRepairSiteCode)) Then
	'
	' Create the "RepairOutput" folder
	'
	g_strRepairOutputPath = g_strParentFolder & "RepairOutput_ClientHealth\"
	Call g_objFunctions.BuildPath(g_strRepairOutputPath)
	g_rsMachinesToProcess.MoveFirst
	While Not g_rsMachinesToProcess.EOF
		g_strMachine = g_rsMachinesToProcess("Machine")
		'
		' Setup the name of the logging file
		'
		g_strLoggingFile = g_strMachine
		'
		' Setup Logging and Tracing files
		'
		g_strGUID = g_objFunctions.CreateGloballyUniqueID()
		g_blnCreateNewFile = True
		g_strLogAndTrace = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_FILE
		g_strLogAndTraceErrors = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_ERROR_FILE
		g_strLogAndTraceClientResolution = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_CLIENT_RESOLUTION_FILE
		g_strLogAndTracePossibleProcessing = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_POSSIBLE_PROCESSING_FILE
		g_strLogAndTraceRegistryProcessing = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_REGISTRY_PROCESSING_FILE
		g_strLogAndTraceExecCmdGeneric = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_EXEC_CMD_GENERIC_FILE
		g_strLogAndTraceLoadRS = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_LOADRS_FILE
		g_strLogAndTraceXMLFileAndRegistry = g_strTraceOutputPath & g_strLoggingFile & "_" & CStr(g_strGUID) & "_" & CLIENTHEALTH_TRACE_XML_FILE_AND_REGISTRY
		'
		' Setup global access to local function logging classes
		'
		Set g_objLogAndTrace = (New LoggingAndTracing)(g_strLogAndTrace, g_blnTraceBasic, g_blnCreateNewFile)
		Set g_objLogAndTraceErrors = (New LoggingAndTracing)(g_strLogAndTraceErrors, True, g_blnCreateNewFile)
		Set g_objLogAndTraceClientResolution = (New LoggingAndTracing)(g_strLogAndTraceClientResolution, g_blnTraceClientResolution, g_blnCreateNewFile)
		Set g_objLogAndTracePossibleProcessing = (New LoggingAndTracing)(g_strLogAndTracePossibleProcessing, g_blnTracePossibleProcessing, g_blnCreateNewFile)
		Set g_objLogAndTraceRegistryProcessing = (New LoggingAndTracing)(g_strLogAndTraceRegistryProcessing, g_blnTraceRegistryProcessing, g_blnCreateNewFile)
		Set g_objLogAndTraceExecCmdGeneric = (New LoggingAndTracing)(g_strLogAndTraceExecCmdGeneric, g_blnTraceExecCmdGeneric, g_blnCreateNewFile)
		Set g_objLogAndTraceLoadRS = (New LoggingAndTracing)(g_strLogAndTraceLoadRS, g_blnTraceLoadRS, g_blnCreateNewFile)
		Set g_objLogAndTraceXMLFileAndRegistry = (New LoggingAndTracing)(g_strLogAndTraceXMLFileAndRegistry, g_blnTraceXMLFileAndRegistry, g_blnCreateNewFile)
		Call DoRepairProcessing(g_strMachine, g_intFlag)
		g_objFunctions.DeleteAllRecordsetRows(g_rsGeneric)
		g_objFunctions.DeleteAllRecordsetRows(g_rsConnectionStatus)
		g_objFunctions.DeleteAllRecordsetRows(g_rsClientConfiguration)
		'
		' Clean up logging and tracing in preparation for next machine
		'
		Set g_objLogAndTrace = Nothing
		Set g_objLogAndTraceErrors = Nothing
		Set g_objLogAndTraceClientResolution = Nothing
		Set g_objLogAndTracePossibleProcessing = Nothing
		Set g_objLogAndTraceRegistryProcessing = Nothing
		Set g_objLogAndTraceExecCmdGeneric = Nothing
		Set g_objLogAndTraceLoadRS = Nothing
		Set g_objLogAndTraceXMLFileAndRegistry = Nothing
		'
		' Process next machine
		'
		g_rsMachinesToProcess.MoveNext
	Wend
End If
'
' Cleanup
'
Set g_objFSO = Nothing
Set g_objFunctions = Nothing
Set g_objXMLProcessing = Nothing
Set g_objPossibleProcessing = Nothing
Set g_objClientResolution = Nothing
Set g_objRegistryProcessing = Nothing
Set g_rsMachinesToProcess = Nothing
Set g_rsGeneric = Nothing
Set g_rsClientConfiguration = Nothing
Set g_rsConnectionStatus = Nothing
Set g_objSourceFile = Nothing
Set g_objFolder = Nothing
Set g_colFiles = Nothing
