	
'''##################################################################################################################################################################
''#		Script Name: 		LE_Seeker_OMJ_Verify_Resume																												#
''#		Script Owner: 		Srini																																	#	
''#		Date of creation: 	09/20/2019																																#
''#		Functionality:		Verifying OMJ details																													#
''#		Input data:			random Seeker id, OWCMS_LE_Info.xls file under C:\UFT_TestData\LE_Input_Files															#
'''##################################################################################################################################################################

'''AAAAAAAAAAAACCCCCCCAAAAAAAAAAAAAABBBBBBB
'''Loading Function libraries
sFunctionLibraryDir = Environment.Value("TestDir") + "\..\Function_libraries\"
LoadFunctionLibrary	sFunctionLibraryDir + "CommonActionsLibraryLE.qfl"
LoadFunctionLibrary	sFunctionLibraryDir + "CommonFunctionLibLE.qfl"
LoadFunctionLibrary	sFunctionLibraryDir + "SignInPageLE.qfl"


''' Reading INPUT file
DataFilePath =  Environment.Value("TestDir") + "\Input Files\"		''' Input File Location
Call ReadExcel_AUT("OWCMS_LE_Info.xls")				'''reading input Data file

'''To start output after a break line
print clear

'''Closing all browsers, Excel files open
KillProcess				

'''Temp Variables
Dim conObj, rsObj, sqlSTMT, strSheetName
Dim intSeekerId, strResumeId, strResumeTitle, intResumeActiveStatus, intRowInd

'''Connecting DB getting seeker id
set conObj = CreateObject("ADODB.Connection")
set rsObj = CreateObject("ADODB.Recordset")

'''To get the Action name to prepare Output sheet
Set qtApp = CreateObject("QuickTest.Application")
'MsgBox qtApp.Test.Actions(1).Name
strSheetName = qtApp.Test.Actions(1).Name


'''OMJ needs UAT DB''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''			
strConnectionString ="Provider=OraOLEDB.Oracle; Data Source=" & _
			"(DESCRIPTION=(CID=GTU_APP)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & "owcms-nprd-scan" _
			& ")(PORT="& "1521" &")))(CONNECT_DATA=(SERVICE_NAME="&"SCTU"&")(SERVER=DEDICATED)));" & _
			"User Id="& "SCOTI_RESEARCH" &";Password="& "SCOTI_RESEARCH" &";"			
conObj.Open strConnectionString


'''Query to get unique Seeker id based on LE, and have valid Resumes 		
sqlSTMT = "SELECT * FROM ( " _
				& "SELECT B.SEEKER_ID, C.FIRST_NAME, C.LAST_NAME , A.MGS_RESUME_ID, A.RESUME_TITLE, A.ACTIVE_RESUME_FLAG " _
				& "FROM EOMJ_RESUME_DATA A,  SEEKER_CASE_DATA B, SEEKER_DATA C " _
				& "WHERE A.SEEKER_ID = B.SEEKER_ID " _
				& "AND B.SEEKER_ID = C.SEEKER_ID " _
				& "AND B.SEEKER_TYPE ='LE' order by dbms_random.value) tbl " _
				& "where tbl.MGS_RESUME_ID IS NOT NULL and rownum = 1"

rsObj.Open sqlSTMT,conObj

intSeekerId	=	rsObj("SEEKER_ID").value						''' getting Seeker ID from DB
strFirstName = rsObj("FIRST_NAME").value						''' getting Seeker First Name
strLastName = rsObj("LAST_NAME").value						''' getting Seeker First Name
strResumeId = rsObj("MGS_RESUME_ID").value							'''' getting Resume ID from DB
strResumeTitle = rsObj("RESUME_TITLE").value						'''' getting Resume title from DB
intResumeActiveStatus = rsObj("ACTIVE_RESUME_FLAG").value 			'''' getting Resume status flag from DB
		
conObj.Close
Set conObj = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''launching OWCMS AUT(ApplicationUnderTest)
Call StartSystemApp(strBrowser,strURL)	

'''Verifying SignIn''''
Call SystemSignIn (cstr(strSysUID),cstr(strSysPWD))		

Call VerifyAnyWebElementItemExist ("LABEL", "", "Current Office Assigned")												''''Verifying Office names window
If intErrorcode=0 Then
	Print getLogTime() &  CommonFunctionLibPage + "Office Names window not appeared"									''''Writing step info to log
	Call CaptureImage ("Current Office Assigned")																		''''to save error screen shot page
	ExitTest																											''''To exit test case execution
else
		Print getLogTime() &  CommonFunctionLibPage + "Office Names: appeared"											''''Writing step info to log
End If
	

Call VerifyAnyWebElementItemExist ("SPAN", "lovFormChooseOffice:lovDataTableChooseOffice:125:officeName" , strOfficeName)	'''Verifying Office name
If intErrorcode=0 Then
	Print getLogTime() &  CommonFunctionLibPage + "Office Name is not valid"											''''Writing step info to log
	Call CaptureImage ("Office Name is not valid")																		''''to save error screen shot page
	ExitTest																											''''To exit test case execution
else
	Call selectOffice("SPAN", strOfficeName)
	Print getLogTime() &  CommonFunctionLibPage + "Office Name is valid"
End If


Call ClickOnLinkOrButton("lovFormChooseOffice:lovSubmitButtonChooseOffice","input", "OK", "button", "OK", "//INPUT[@id=""lovFormChooseOffice:lovSubmitButtonChooseOffice""]")	'''Selecting Office name
If intErrorcode=0 Then	
	Print getLogTime() &  CommonFunctionLibPage + "Error in Selecting Office"											''''Writing step info to log
	Call CaptureImage ("Error in Selecting Office")																		''''to save error screen shot page
	ExitTest																											''''To exit test case execution
else
	Print getLogTime() &  CommonFunctionLibPage + strOfficeName + " :Office selected"	
End If


Call VerifyAnyWebElementItemExist ("DIV", "" , strMenuToVerify)															''''Verifying Menu 
If intErrorcode=0 Then
	Print getLogTime() &  CommonFunctionLibPage + strMenuToVerify + " menu is not there"								''''Writing step info to log
	Call CaptureImage ("Error in Selecting menu")																		''''to save error screen shot page
	ExitTest
else
	Print getLogTime() &  CommonFunctionLibPage + strMenuToVerify + " menu available"									''''Writing step info to log
End If
 

 	'''Adding EXCEL OUTPUT FILe headers''''''''''''''''''''''
	ITVariable  = DataTable.GetSheet(strSheetName).AddParameter("SEEKERID","")  										'''' Preparing Seeker Id for OUTPUT, COLUMN 1
	ITVariable  = DataTable.GetSheet(strSheetName).AddParameter("FIRST_NAME","")  										'''' Preparing Seeker Id for OUTPUT, COLUMN 2
	ITVariable  = DataTable.GetSheet(strSheetName).AddParameter("LAST_NAME","")  										'''' Preparing Seeker Id for OUTPUT, COLUMN 3
	ITVariable  = DataTable.GetSheet(strSheetName).AddParameter("MGS_RESUME_ID", "")									'''' Preparing RESUME ID for OUTPUT, column 4
	ITVariable  = DataTable.GetSheet(strSheetName).AddParameter("RESUME_TITLE", "")										'''' Preparing RESUME TITLE for OUTPUT, column 5
	ITVariable  = DataTable.GetSheet(strSheetName).AddParameter("RESUME_STATUS", "")										'''' Preparing RESUME STATUS for OUTPUT, column 6
	
 
	call ClickMenu("A","Select Job Seeker","Select Job Seeker", "//DIV[@id=""menu""]/UL[1]/LI[1]/UL[1]/LI[7]/A[1]")		''''Verifying valid submenu
	If intErrorcode=0 Then
		Print getLogTime() &  CommonFunctionLibPage + " Select Job Seeker menu is not there"							''''Writing step info to log
		Call CaptureImage ("Error in Selecting menu")																	''''to save error screen shot page
		ExitTest
	else
		Print getLogTime() &  CommonFunctionLibPage + "Select Job Seeker menu available"	
	End If
	
	'searchSeeker(strFirstName-1,strLastName-2, strDOB-3, strGender-4, intSSN-5, intSEEKERID-6, strEmail-7, intZIP-8)	PARAMETERS list
	Call searchSeeker("", "", "", "", "", intSeekerId, "", "")													'''' Searching Seeker
				
	Call ClickOnLinkOrButton("idLocateSeekerGrid:0:j_id.*","INPUT","idLocateSeekerGrid:0:j_id.*","button", "","")		'''' opening Seeker info
	Wait 1
	Call ClickOnLinkOrButton("idCaseHistoryGrid:0:j_id134","INPUT","idCaseHistoryGrid:0:j_id134","button", "","")		'''' opening Seeker info
	wait 2
	call ClickMenu("A","OMJ Details","OMJ Details", "//DIV[@id=""menu""]/UL[1]/LI[1]/UL[1]/LI[14]/A[1]")				'''' Opening OMJ menu details
	wait 1
	
	'Call VerifyAnyWebElementItemExist ("SPAN", "idEomjResumeData:0:idResumeTitle" ,strResumeTitle)						'''''comparing and highlightinh matching Ttile
	'Call VerifyAnyWebElementItemExist ("SPAN", "idEomjResumeData:0:idResumeTitle" ,strResumeId)							'''''comparing and highlightinh matching ID
	
	Call VerifyAnyWebElementItemExist ("SPAN", "idEomjResumeData:.:idResumeId" ,strResumeId)							'''''comparing and highlightinh matching ID	
	If intErrorcode=0 Then
		Print getLogTime() &  strResumeId & " : Resume Id Not found"	
		ExitTest
	else
		Print getLogTime() &  strResumeId & " : Resume Id Found"							''''Writing step info to log
		startPosition = InStr(strIDCtrlName, ":")
		endPosition = InStrRev(strIDCtrlName, ":")
		intRowInd = Mid(strIDCtrlName, (startPosition+1), (endPosition - (startPosition+1)))
	End If

	Call VerifyAnyWebElementItemExist ("SPAN", "idEomjResumeData:" & intRowInd & ":idResumeTitle" ,strResumeTitle)						'''''comparing and highlightinh matching Ttile
	If intErrorcode=0 Then
		Print getLogTime() & strResumeTitle & " : Resume Title NOT found"
		ExitTest		
	else
		Print getLogTime() &  strResumeTitle & " : Resume Title found"							''''Writing step info to log
	End If


	DataTable.GetSheet(strSheetName).GetParameter("SEEKERID").Value = intSeekerId										'''' Preparing SEEKER ID for OUTPUT
	DataTable.GetSheet(strSheetName).GetParameter("FIRST_NAME").Value = strFirstName									'''' Preparing FIRST NAME for OUTPUT 	
	DataTable.GetSheet(strSheetName).GetParameter("LAST_NAME").Value = strLastName										'''' Preparing LAST NAME for OUTPUT 	
	DataTable.GetSheet(strSheetName).GetParameter("MGS_RESUME_ID").Value = strResumeId									'''' Preparing RESUME ID for OUTPUT 	
	DataTable.GetSheet(strSheetName).GetParameter("RESUME_TITLE").Value = strResumeTitle								'''' Preparing RESUME TITLE for OUTPUT 	
	
	if intResumeActiveStatus = 1 then
		DataTable.GetSheet(strSheetName).GetParameter("RESUME_STATUS").Value = "Active"									'''' Preparing Resume status for OUTPUT 
	else
		DataTable.GetSheet(strSheetName).GetParameter("RESUME_STATUS").Value = "InActive"								'''' Preparing Resume status for OUTPUT 
	End If
		
	''' writing to OUTPUT file
	Dim dt : dt = now()
	sTimeStamp = sprintf("{0:yyyyMMdd_HHmmss}", Array(dt))
	OutputFile = Environment.Value("TestDir") + "/Output files/Results_" + sTimeStamp + ".xls"
 	DataTable.ExportSheet OutputFile, strSheetName													

 	'''Logging out of OWCMS 
	Call logOutOWCMS("Image", "INPUT", "ribbonExit", "ribbonExit", "", "")			
	Call logOutOWCMS("Image", "INPUT", "ribbonExit", "ribbonExit", "", "")
	Call logOutOWCMS("Image", "INPUT", "ribbonExit", "ribbonExit", "", "")
	Call logOutOWCMS("Image", "INPUT", "ribbonExit", "ribbonExit", "", "")

	KillProcess							'''killing all process left for Chrome, IE, Excel

