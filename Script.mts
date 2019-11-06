'''##############################################################################################################################################################
''#		Script Name		: 		LE_SEEKER																														#
''#		Description		:      	Creating a new LE seeker																										#
''#		Script Owner	: 		JYOTHSNA DUSSA																													#	
''#		Date of creation: 		09/06/2019																														#
''#		Functionality	:		Creating LE_Seeker from OWCMS application.																						#
''#		Input data		:		OWCMS_LE_Info.xls file under C:\UFT_TestData\LE_Input_Files\																	#
' #		Modified By		:      	Jyothsna Dussa,  Srinivasa Nagireddy																							#
' #		Modified Date	:    	09/16/2019, 10/28/2019																											#
' #		Purpose of Modification: Updating as per the Coding Standards																							#
' #								Fixed the Dynamic name id's and updated the script according to Data input format 												#
'''##############################################################################################################################################################

'''Loading Function libraries dynamically
LoadFunctionLibrary	Environment.Value("TestDir") + "\Function_libraries\CommonActionsLibraryLE.qfl"
LoadFunctionLibrary	Environment.Value("TestDir") + "\Function_libraries\CommonFunctionLibLE.qfl"
LoadFunctionLibrary	Environment.Value("TestDir") + "\Function_libraries\SignInPageLE.qfl"


print clear					'''To start output after a break line

KillProcess					'''Closing all browsers

Reporter.Filter = 2 

'''Temp Variables
Dim conObj, rsObj, sqlSTMT, strSheetName
Dim nUsedRows, nUsedCols, nRow, nCol
Dim intSeekerId, strResumeId, strResumeTitle, intResumeActiveStatus

''' Reading INPUT file
DataFilePath =  "C:\UFT_TestData\LE_Input_Files\"		''' Input File Location
Call ReadExcel_AUT("OWCMS_LE_Info.xls")				'''reading input Data file

''''To get the Action name to prepare Output sheet
'Set qtApp = CreateObject("QuickTest.Application")
'strSheetName = qtApp.Test.Actions(1).Name


Call StartSystemApp(strBrowser,strURL)					'''launching OWCMS AUT(ApplicationUnderTest)

Call SystemSignIn (cstr(strSysUID),cstr(strSysPWD))		'''Verifying SignIn													

Call VerifyAnyWebElementItemExist ("LABEL", "", "Current Office Assigned")				''' Verifying Office window
If intErrorcode=0 Then
	Print getLogTime() &  CommonFunctionLibPage + "County Names window not appeared" 	'''Writing step info to log
	Call CaptureImage ("Current Office Assigned")									 	'''To save error screen shot page
	ExitTest																		 	'''To exit test case execution
else
		Print getLogTime() &  CommonFunctionLibPage + "County Names: appeared"	     	'''Writing step info to log
End If
	

Call VerifyAnyWebElementItemExist ("SPAN", "lovFormChooseOffice:lovDataTableChooseOffice:125:officeName" , strOfficeName)		''' Verifying, and Launching Office name
If intErrorcode=0 Then
	Print getLogTime() &  CommonFunctionLibPage + "Office Name is not valid"		 	'''Writing step info to log
	Call CaptureImage ("County Name is not valid")									 	'''to save error screen shot page
	ExitTest																		 	'''To exit test case execution
else
	Call selectOffice("SPAN", strOfficeName)											''' Selecting OFFICE
	Print getLogTime() &  CommonFunctionLibPage + "Office Name is valid"				'''Writing step info to log
End If


Call ClickOnLinkOrButton("lovFormChooseOffice:lovSubmitButtonChooseOffice", "INPUT", "OK", "button", "OK", "//INPUT[@id=""lovFormChooseOffice:lovSubmitButtonChooseOffice""]")	'''Verifying Launching Office 
If intErrorcode=0 Then
	Print getLogTime() &  CommonFunctionLibPage + "Error in Selecting County/Office"	'''Writing step info to log
	Call CaptureImage ("Error in Selecting County")										'''to save error screen shot page
	ExitTest																			'''To exit test case execution
else
	Print getLogTime() &  CommonFunctionLibPage + strOfficeName + " :Office selected"	'''Writing step info to log
End If


Call VerifyAnyWebElementItemExist ("DIV", "",  strMenuToVerify)							''' verifyingf Menu name
If intErrorcode=0 Then
	Print getLogTime() &  CommonFunctionLibPage + strMenuToVerify + " menu is not there" '''Writing step info to log
	Call CaptureImage ("Error in Selecting menu")								           '''to save error screen shot page
	ExitTest
else
	Print getLogTime() &  CommonFunctionLibPage + strMenuToVerify + " menu available"		'''Writing step info to log
End If
 

''' Adding OUTPUT FILe header
DataTable.AddSheet("Results") 
ITVariable  = DataTable.GetSheet("Results").AddParameter("SNO","")  			''' Preparing SNO for OUTPUT, COLUMN 1

ITVariable  = DataTable.GetSheet("Results").AddParameter("FIRSTNAME","") 		''' Preparing for OUTPUT, COLUMN 2

ITVariable  = DataTable.GetSheet("Results").AddParameter("LASTNAME","")  		''' Preparing for OUTPUT, COLUMN 3

ITVariable  = DataTable.GetSheet("Results").AddParameter("SEEKERID","")  		''' Preparing Seeker Id for OUTPUT, COLUMN 4

ITVariable  = DataTable.GetSheet("Results").AddParameter("RECORD_STATUS", "") 	''' Adding column for record status, , COLUMN 5
 

nUsedRows = datatable.getsheet("Seeker_Data").getrowcount()			''' Get the number of used rows

nUsedCols = datatable.getsheet("Seeker_Data").GetParameterCount()	''' Get the number of used columns

DataTable.SetCurrentRow(1)

' Loop through each column
For nRow = 1 To nUsedRows
		
		' Exiting when SNO column is blank
    	If datatable.getsheet("Seeker_Data").GetParameter("SNO").Value= ""  Then
    		Exit for
    	End If


		call ClickMenu("A","Select Job Seeker","Select Job Seeker", "//DIV[@id=""menu""]/UL[1]/LI[1]/UL[1]/LI[7]/A[1]")		''''Verifying valid submenu
		If intErrorcode=0 Then
			Print getLogTime() &  CommonFunctionLibPage + " Select Job Seeker menu is not there"							''''Writing step info to log
			Call CaptureImage ("Error in Selecting menu")																	''''to save error screen shot page
			ExitTest
		else
			Print getLogTime() &  CommonFunctionLibPage + "Select Job Seeker menu available"								''''Writing step info to log
		End If
	
		
		''' Entering Search criteria
		'searchSeeker(strSheetName, strFirstName,strLastName, DOB,Gender, intSSN<from DB only>,SeekerID, strEmail<currently NOT ABLE TO USE now, will correct it>, intZIP)
		Call searchSeeker(DataTable.GetSheet("Seeker_Data").GetParameter("First_Name").Value, DataTable.GetSheet("Seeker_Data").GetParameter("Last_Name").Value, _
		DataTable.GetSheet("Seeker_Data").GetParameter("Date_of_Birth").Value, DataTable.GetSheet("Seeker_Data").GetParameter("Gender").Value, _
		DataTable.GetSheet("Seeker_Data").GetParameter("SSN").Value, DataTable.GetSheet("Seeker_Data").GetParameter("Seeker_Id").Value, _
		"", DataTable.GetSheet("Seeker_Data").GetParameter("Street_Zip_Code").Value)
				
		''' Verifying error message for search			'1 Means displayed Results,0 Means  error displayed and No Results
		Call VerifyStatusMessage("Seeker record not found. Please try searching with another combination: - Email Address - First and Last Name - SSN - Seeker Id - Zip Code with First and Last Name")		
		
		'Entering data as No matching Seeker
		If intErrorcode = 1 Then	
			
			Call ClickOnLinkOrButton("idNewSeeker", "INPUT", "New Seeker", "button", "New Seeker", "//INPUT[@id=""idNewSeeker""]")
			
			Call enterData("WebEdit","INPUT", "idssn", "ssn", "text",DataTable.GetSheet("Seeker_Data").GetParameter("SSN").Value)							'''entering SSN
			
			Call enterData("WebEdit","INPUT", "streetZip", "streetZip", "text", DataTable.GetSheet("Seeker_Data").GetParameter("Street_Zip_Code").Value)	'''entering Zip code	
			
			call ClickOnLinkOrButton("idCONTINUE", "INPUT","Continue","button", "Continue","//INPUT[@id=""idCONTINUE""]")									'''clikcing Continue
			
			Call enterData("WebEdit","INPUT", "idIntakeDate", "idIntakeDate", "text", DataTable.GetSheet("Seeker_Data").GetParameter("Intake_Date").Value) 	'''Entering INTAKE date	
			
			Call enterData("WebEdit","INPUT", "idAddress1", "idAddress1", "text",DataTable.GetSheet("Seeker_Data").GetParameter("Street_Address").Value) 	''' ENTERIING ADDRESS 1
			
			Call enterData("WebEdit","INPUT", "idEmail", "idEmail", "text", DataTable.GetSheet("Seeker_Data").GetParameter("Email_Address").Value)			''' ENTERING EMAIL ADDRESS
			
		'Entering data as Seeker Matching search results
		ElseIf intErrorcode = 0 Then																			
			
			call ClickOnLinkOrButton("idCreateNewSeeker", "INPUT","Create New Seeker","button", "Create New Seeker","//INPUT[@id=""idCreateNewSeeker""]")	'''Entering Seeker data if already exists
			
			Call enterData("WebEdit","INPUT", "idIntakeDate", "idIntakeDate", "text", DataTable.GetSheet("Seeker_Data").GetParameter("Intake_Date").Value) 	''' Entering INTAKE date	
			
			Print getLogTime() &  CommonFunctionLibPage + "Creating a new seeker process"																	'''Writing step info to log
			
			intErrorcode = 1			''' Restting Error code to 1
		End If				

		Call enterData("WebEdit","INPUT", "idDateOfBirth", "idDateOfBirth", "text", DataTable.GetSheet("Seeker_Data").GetParameter("Date_of_Birth").Value)	''' enter DOB 
		
		Call enterData("WebList","SELECT", "idGender", "idGender", "", DataTable.GetSheet("Seeker_Data").GetParameter("Gender").Value)						''' Selecting Gender		
		'''Verifying Gender selection
		If intErrorcode=0 Then
			Print getLogTime() &  CommonFunctionLibPage + "GENDER is not selected"		''' Writing step info to log
			Call CaptureImage ("Error in Selecting menu")								'''to save error screen shot page
			ExitTest
		else
			Print getLogTime() &  CommonFunctionLibPage + "selected valid Gender"		''' Writing step info to log
		End If
		
		Call enterData("WebList","SELECT", "idEthnicity", "idEthnicity", "", DataTable.GetSheet("Seeker_Data").GetParameter("Ethnicity").Value)				''' Selecting Ethnitcity	
					
		Call enterData("WebList","SELECT", "idStaffcitizenship", "idStaffcitizenship", "", DataTable.GetSheet("Seeker_Data").GetParameter("Citizenship").Value)	''' Selecting CITIZENSHIP
		
		call enterData("Image","INPUT", "idRaceInsert", "idRaceInsert", "","")				''' Entering RACE

		Call enterData("WebList","SELECT", "idRaceList:0:idRace", "idRaceList:0:idRace", "", DataTable.GetSheet("Seeker_Data").GetParameter("Race").Value) 		''' Selecting Ethnitcity
		
		Call enterData("WebRadioGroup","INPUT","idEmplStatus:.*","idEmplStatus","radio", DataTable.GetSheet("Seeker_Data").GetParameter("idEmplStatus").Value)	''' Selecting Employee status


		'''To activate Additoinal label TAB, as some action is needed(TEMP) 
'		call enterData("Image","INPUT", "ribbonSave", "ribbonSave", "","")								''' Button to click to save General tab	
'		Call ClickOnLinkOrButton("", "INPUT", "Yes", "button", "Yes", "//SPAN[@id=""dialogchangeConfirmation""]/DIV[1]/INPUT[1]") ' Button to click to save on popup	
		call enterData("WebEdit","INPUT", "idPrimaryPhone", "idPrimaryPhone", "text","(544) 657-4567")
		Call enterData("WebEdit","INPUT", "idIntakeDate", "idIntakeDate", "text", DataTable.GetSheet("Seeker_Data").GetParameter("Intake_Date").Value) 			''' Entering INTAKE date	


		Call enterdata("WebElement","TD","idadditional_lbl","Additional","","//TD[@id=""idadditional_lbl""]")													''' opening ADDITIONAL tab


		Call enterData("WebList","SELECT", "idEducationlevel", "idEducationlevel", "", DataTable.GetSheet("Seeker_Data").GetParameter("Education_Level").Value)		''' Selecting Educationlevel
				
		Call enterData("WebList","SELECT", "idEducationStatus", "idEducationStatus", "", DataTable.GetSheet("Seeker_Data").GetParameter("Education_Status").Value) 								''' Selecting Education status
			
		Call enterData("WebRadioGroup","INPUT", "idMilitaryFlag:.", "idMilitaryFlag", "radio",DataTable.GetSheet("Seeker_Data").GetParameter("Have_you_servered_in_US_Military").Value)			''' Selecting Have you served in US Military
						
		Call enterData("WebRadioGroup","INPUT", "idVetSpouseFlag:.", "idVetSpouseFlag", "radio",DataTable.GetSheet("Seeker_Data").GetParameter("Are_you_a_Spouse_of_Veteran").Value)			''' Selecting Are you a Spouse of Veteran?\

		Call enterData("WebRadioGroup","INPUT", "idHomeLessRadio:.", "idHomeLessRadio", "radio",DataTable.GetSheet("Seeker_Data").GetParameter("Homeless_veteran").Value)						''' Selecting Homeless Veteran?
					
		Call enterdata("WebElement","TD","idProgramData2_lbl","Program Data","","//TD[@id=""idProgramData2_lbl""]")							''' opening PROGRAM DATA LABEL		
		
		Call enterData("WebRadioGroup","INPUT", "idLESelectiveServiceFlag:*.", "idLESelectiveServiceFlag", "radio", DataTable.GetSheet("Seeker_Data").GetParameter("Registered_for_Selective_Service").Value)		''' Selecting SelectiveServiceFlag	


'Browser("OWCMS - Select Seeker").Page("OWCMS - Basic Intake_2").WebRadioGroup("j_id997").Select "0"							''''********** RECORDED LINE******************************************************************** @@ hightlight id_;_Browser("OWCMS - Select Seeker").Page("OWCMS - Basic Intake 2").WebRadioGroup("j id997")_;_script infofile_;_ZIP::ssf7.xml_;_
'		'Call enterData("WebRadioGroup","INPUT", "j_id.*:1", "j_id.*", "radio",DataTable.GetSheet("Seeker_Data").GetParameter("SNAP_Employment_Training_is_required").Value)						''' Selecting Homeless Veteran?
''		Call enterData("WebRadioGroup","INPUT", "j_id998:1", "j_id998", "radio",DataTable.GetSheet("Seeker_Data").GetParameter("SNAP_Employment_Training_is_required").Value)						''' Selecting Homeless Veteran?
'		Call enterData("WebRadioGroup","INPUT", "j_id.*", "j_id.*", "radio",DataTable.GetSheet("Seeker_Data").GetParameter("SNAP_Employment_Training_is_required").Value)						''' Selecting Homeless Veteran?
''		Call enterData("WebRadioGroup","INPUT", "j_id.*:.", "j_id.*", "radio",DataTable.GetSheet("Seeker_Data").GetParameter("SNAP_Employment_Training_is_required").Value)						''' Selecting Homeless Veteran?
		
			Set inputCtrl = Description.Create
			inputCtrl("micclass").Value = "WebRadioGroup"
			inputCtrl("type").Value = "radio"
			inputCtrl("html id").RegularExpression = True
			inputCtrl("html id").Value = "j_id.*:."
			inputCtrl("name").RegularExpression = True
			inputCtrl("name").Value = "j_id.*"
			inputCtrl("xpath").RegularExpression = True
			inputCtrl("xpath").Value = "//TABLE[@id=""idCoreServiceTab""]/TBODY[1]/TR[2]/TD[1]/TABLE[1]/TBODY[1]/TR[1]/TD[1]/FIELDSET[1]/TABLE[7]/TBODY[1]/TR[2]/TD[5]/TABLE[1]/TBODY[1]/TR[1]/TD[2]/INPUT[1]"			
			XPage.WebRadioGroup(inputCtrl).select "1"

			
		call enterData("Image","INPUT", "ribbonSave", "ribbonSave", "","")																	''' Button to click to save General tab	

		Call ClickOnLinkOrButton("", "INPUT", "Yes", "button","Yes", "//SPAN[@id=""dialogchangeConfirmation""]/DIV[1]/INPUT[1]")			'''Button to click to save on popup
				
		DataTable.GetSheet ("Results").SetCurrentRow (nrow)  			''' setting cursor in next row
				
		' Adding First name, Last name to output
		DataTable.GetSheet("Results").GetParameter("SNO").Value = nRow																			''' Preparing Serial number for OUTPUT 
		DataTable.GetSheet("Results").GetParameter("FIRSTNAME").Value =  DataTable.GetSheet("Seeker_Data").GetParameter("First_Name").Value		''' Preparing SEEKER First name for OUTPUT 
		DataTable.GetSheet("Results").GetParameter("LASTNAME").Value = DataTable.GetSheet("Seeker_Data").GetParameter("Last_Name").Value		''' Preparing SEEKER Last name for OUTPUT 
				
		' if creation is success or Not
		Call VerifySave ("Record(s) saved.")	
		
		''' ADDING SEEKER ID FROM DB
		If intErrorcode = 1 Then

			set conObj = CreateObject("ADODB.Connection")
			set rsObj = CreateObject("ADODB.Recordset")
			
			connectionString ="Provider=OraOLEDB.Oracle; Data Source=" & _
			"(DESCRIPTION=(CID=GTU_APP)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & "owcms-nprd-scan" & ")(PORT="& "1521" &")))(CONNECT_DATA=(SERVICE_NAME="&"SNTS"&")(SERVER=DEDICATED)));" & _
			"User Id="& "SCOTI_RESEARCH" &";Password="& "SCOTI_RESEARCH" &";"
			
			conObj.Open connectionString
			
			sqlSTMT = "SELECT SEEKER_ID FROM SEEKER_DATA A WHERE  FIRST_NAME = '" & ucase(DataTable.GetSheet("Seeker_Data").GetParameter("First_Name").Value) & "' AND " & "LAST_NAME = '" & _
			ucase(DataTable.GetSheet("Seeker_Data").GetParameter("Last_Name").Value) & "' order by SEEKER_ID desc fetch first 1 row only"
			
			rsObj.Open sqlSTMT,conObj
						
			DataTable.GetSheet("Results").GetParameter("SEEKERID").Value = rsObj("SEEKER_ID").value				''' Preparing SEEKER ID for OUTPUT 
			
			DataTable.GetSheet("Results").GetParameter("RECORD_STATUS").Value = "Record created sucessfully: "+ getLogTime	''' Adding record status
					
			conObj.Close
			Set conObj = Nothing

		Else
		
			DataTable.GetSheet("Results").GetParameter("RECORD_STATUS").Value = "Fail, Error in record creation: "+ getLogTime	''' Adding record status
			
			Call ClickOnLinkOrButton("", "input", "No", "button", "No", "//SPAN[@id=""menuForm:dialogmenuExitConfirmation""]/DIV[1]/INPUT[2]")			'''Button to click to save on popup
			
		End If
    
		print clear						'''	Printing Line break
		
		DataTable.GetSheet("Seeker_Data").SetNextRow		''' Moving cursor to next row
	
 	Next


	'''Deleteing if file exists
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(DataFilePath + "Results.xls") Then
        fso.DeleteFile(DataFilePath + "Results.xls")
    End If
    set fso = nothing
 	'DataTable.ExportSheet DataFilePath + "Results.xls", "Results", strSheetName									'''Generating Output sheet
 	DataTable.ExportSheet DataFilePath + "Results.xls", "Results"									'''Generating Output sheet

''''Logging out of OWCMS 
Call logOutOWCMS("Image", "INPUT", "ribbonExit", "ribbonExit", "", "")			
Call logOutOWCMS("Image", "INPUT", "ribbonExit", "ribbonExit", "", "")
Call logOutOWCMS("Image", "INPUT", "ribbonExit", "ribbonExit", "", "")
Call logOutOWCMS("Image", "INPUT", "ribbonExit", "ribbonExit", "", "")

KillProcess 			'''	killing all browwsers, Excel files opened

