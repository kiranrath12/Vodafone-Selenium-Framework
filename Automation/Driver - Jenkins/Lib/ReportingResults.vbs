

'#################################################################################################
'         Function Name : CreateAXLSResultfile
'         Description : This function creates a .xls result file to store Test Case Id and Result
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function CreateAXLSResultfile()
	sPathExcel = Environment.Value("ResultPath") & ".xls"
	sTimeStamp=TimeStamp_fn()
	Set sObjExcel = CreateObject("Excel.Application")
	Set sObjWbk = sObjExcel.Workbooks.open (sPathExcel)
	Set sObjSheet = sObjExcel.ActiveWorkbook.Worksheets(1)
	sObjWbk.SaveAs  Environment.Value("ResultPath") & sTimeStamp & ".xls" 
	Environment.Value("SavedResultPath") = Environment.Value("ResultPath") & sTimeStamp & ".xls" 
	sObjExcel.Quit
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function
'#################################################################################################
'         Function Name : CreateAXMLResultfile
'         Description : This function creates a xml result file to store test case results
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function CreateAXMLResultfile(sPrefixXmlResultFileWith)

	'sTimeStamp=TimeStamp_fn()
'	sResultFileName = sProjectDir & "Driver\Results\" & sProduct_G & "\" & sPrefixXmlResultFileWith  & ".xml"
	sResultFileName = Environment.Value("RunFolderPath") & "\" & sPrefixXmlResultFileWith  & ".xml"
'	Environment.Value("RunFolderPath")
	Dim fso, MyFile
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set MyFile = fso.CreateTextFile(sResultFileName, True)
	MyFile.WriteLine("<?xml version="&chr(34) & "1.0" & chr(34) & " encoding=" & chr(34) & "UTF-8"  & chr(34) & "?>")
	MyFile.WriteLine("<!DOCTYPE RESULT SYSTEM " & chr(34) &  "Result.dtd" & chr(34) & " []>")
	MyFile.WriteLine("<?xml-stylesheet type="  & chr(34) & "text/xsl" &chr(34) & " href=" &chr(34) & "Result.xsl" &chr(34) & "?>")
	MyFile.WriteLine("<RESULT>")
	MyFile.WriteLine("<TITLE>")
	MyFile.WriteLine(left(sPrefixXmlResultFileWith,len(sPrefixXmlResultFileWith)-1))
	MyFile.WriteLine("</TITLE>")
	MyFile.WriteLine("<SUMMARY>")
	MyFile.WriteLine("<TOTAL>")
	MyFile.WriteLine("</TOTAL>")
	MyFile.WriteLine("<TOTALPASS>")
	MyFile.WriteLine("</TOTALPASS>")
	MyFile.WriteLine("<TOTALFAIL>")
	MyFile.WriteLine("</TOTALFAIL>")
	MyFile.WriteLine("<TOTALNOTEXECUTED>")
	MyFile.WriteLine("</TOTALNOTEXECUTED>")
	MyFile.WriteLine("</SUMMARY>")
	MyFile.WriteLine("</RESULT>")
	MyFile.Close
	
	Set oResultdoc = XMLUtil.CreateXML() 
	oResultdoc.LoadFile  sResultFileName
	Set oRoot = oResultdoc.GetRootElement()
	CreateAXMLResultfile=sResultFileName
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function

''#################################################################################################
'         Function Name : CreateTCXMLResultfile
'         Description : This function creates a xml result file to store test case results
'         Created By : 
'         Creation Date :        
'##################################################################################################
REM Public Function CreateTCXMLResultfile(sTestCaseName)

	REM 'sTimeStamp=TimeStamp_fn()
	REM sTCResultFileName = TCFolderPath & "\" & sTestCaseName  & ".xml"
	
	REM Dim fso, MyFile
	REM Set fso = CreateObject("Scripting.FileSystemObject")
	REM Set MyFile = fso.CreateTextFile(sTCResultFileName, True)
	REM MyFile.WriteLine("<?xml version="&chr(34) & "1.0" & chr(34) & " encoding=" & chr(34) & "UTF-8"  & chr(34) & "?>")
	REM MyFile.WriteLine("<!DOCTYPE RESULT SYSTEM " & chr(34) &  "Result.dtd" & chr(34) & " []>")
	REM MyFile.WriteLine("<?xml-stylesheet type="  & chr(34) & "text/xsl" &chr(34) & " href=" &chr(34) & "Result.xsl" &chr(34) & "?>")
	REM MyFile.WriteLine("<RESULT>")
	REM MyFile.WriteLine("<TITLE>")
	REM MyFile.WriteLine(sTestCaseName)
	REM MyFile.WriteLine("</TITLE>")
	REM MyFile.WriteLine("<SUMMARY>")
	REM MyFile.WriteLine("<TOTAL>")
	REM MyFile.WriteLine("</TOTAL>")
	REM MyFile.WriteLine("<TOTALPASS>")
	REM MyFile.WriteLine("</TOTALPASS>")
	REM MyFile.WriteLine("<TOTALFAIL>")
	REM MyFile.WriteLine("</TOTALFAIL>")
	REM MyFile.WriteLine("<TOTALNOTEXECUTED>")
	REM MyFile.WriteLine("</TOTALNOTEXECUTED>")
	REM MyFile.WriteLine("</SUMMARY>")
	REM MyFile.WriteLine("</RESULT>")
	REM MyFile.Close
	
	REM Set oTCResultdoc = XMLUtil.CreateXML() 
	REM oTCResultdoc.LoadFile  sResultFileName
	REM Set oTCRoot = oResultdoc.GetRootElement()
	REM CreateTCXMLResultfile=sTCResultFileName
REM '	If Err.Number <> 0 Then
REM '		iModuleFail = 1
REM '		AddVerificationInfoToResult "Error",Err.Description
REM '	End If
REM End Function


'#################################################################################################
'         Function Name : UpdateTestCaseStatusInResult
'         Description : This function updates test case result in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function UpdateTestCaseStatusInResult(Status)
'	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
'	lastTCNo = TestCases.Count 
'	set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)

    Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE/TCID[text()='"+sTCID+"']/..")
	cnt = TestCases.Count 

	Dim CurrentTC
'	For i=1 to cnt
	Set CurrentTC=TestCases.Item(1)
	CurrentTC.ChildElements().ItemByName("RESULT",1).SetValue(Status)
	 If Status ="PASS" Then
		Call UpdateResultSummaryTotalPass()
	ElseIf Status ="FAIL" Then
		If ExecutionJourney="SecondPart" Then
			Call UpdateResultSummaryTotalFail()
		End IF
	End If
	oResultdoc.SaveFile  sResultFileName
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function

'#################################################################################################
'         Function Name : UpdateTestCaseStatusInResult
'         Description : This function updates test case result in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
'Public Function UpdateTestCaseStatusInResultForSecondPart(Status)
'	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
'	lastTCNo = TestCases.Count 
'    sTCID = adoRS("TCID") & ""
'	Dim CurrentTC
'    For i=1 to lastTCNo
'		Set CurrentTC=TestCases.ItemByName("TESTCASE",i)
'		If  CurrentTC.getValueByXpath("./TCID")= sTCID then
'			CurrentTC.ChildElements().ItemByName("RESULT",1).SetValue(Status)
'			Exit For
'		End If
'	
'	Next
''	set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)
''	CurrentTC.ChildElements().ItemByName("RESULT",1).SetValue(Status)
'	 If Status ="PASS" Then
'		Call UpdateResultSummaryTotalPass()
'	ElseIf Status ="FAIL" Then
'		Call UpdateResultSummaryTotalFail()
'	End If
'	oResultdoc.SaveFile  sResultFileName
''	If Err.Number <> 0 Then
''		iModuleFail = 1
''		AddVerificationInfoToResult "Error",Err.Description
''	End If
'End Function
'
'
''#################################################################################################
''         Function Name : UpdateTestCaseStatusInResultForFirstPart
''         Description : This function updates test case result in the xml result file
''         Created By : 
''         Creation Date :        
''##################################################################################################
'Public Function UpdateTestCaseStatusInResultForFirstPart(Status)
'	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
'	lastTCNo = TestCases.Count 
''    sTCID = adoRS("TCID") & ""
''	Dim CurrentTC
''    For i=1 to cnt
''		Set CurrentTC=TestCases.ItemByName("TESTCASE",i)
''		If  CurrentTC.getValueByXpath("./TCID")= sTCID then
''			CurrentTC.ChildElements().ItemByName("RESULT",1).SetValue(Status)
''			Exit For
''		End If
''	
''	Next
'	set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)
'	CurrentTC.ChildElements().ItemByName("RESULT",1).SetValue(Status)
'	 If Status ="PASS" Then
'		Call UpdateResultSummaryTotalPass()
'	ElseIf Status ="FAIL" Then
'		Call UpdateResultSummaryTotalFail()
'	End If
'	oResultdoc.SaveFile  sResultFileName
''	If Err.Number <> 0 Then
''		iModuleFail = 1
''		AddVerificationInfoToResult "Error",Err.Description
''	End If
'End Function

'#################################################################################################
'         Function Name : UpdateStepStatusInResult
'         Description :  This function updates test case step result in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function UpdateStepStatusInResult(Status)
'	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
'	
'	lastTCNo = TestCases.Count 
'	set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)
	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE/TCID[text()='"+sTCID+"']/..")
	cnt = TestCases.Count 

	Dim CurrentTC
'	For i=1 to cnt
	Set CurrentTC=TestCases.Item(1)
	
	Set Steps = CurrentTC.ChildElements()
	lastStep = Steps.Count 
	
	set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)
	CurrentStep.ChildElements().ItemByName("RESULT",1).SetValue(Status)

	'Adding EXCEL tag for displaying data file path 
'	sExcelName = "C:\Automation\NewCo\Data\" & Environment("sDataFileName") & "#" & Environment("sDataSheetName") & "!a1"
'	CurrentStep.AddChildElementbyname "EXCEL", sExcelName
	
	oResultdoc.SaveFile  sResultFileName

	
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function


'#################################################################################################
'         Function Name : UpdateStepStatusInTCResultFile
'         Description :  This function updates test case step result in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
REM Public Function UpdateStepStatusInTCResultFile(Status)
	REM Set TestCases = oTCResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
	
	REM lastTCNo = TestCases.Count 
	REM set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)
	
	REM Set Steps = CurrentTC.ChildElements()
	REM lastStep = Steps.Count 
	
	REM set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)
	REM CurrentStep.ChildElements().ItemByName("RESULT",1).SetValue(Status)

	REM 'Adding EXCEL tag for displaying data file path 
	REM sExcelName = "C:\Automation\NewCo\Data\" & Environment("sDataFileName") & "#" & Environment("sDataSheetName") & "!a1"
	REM CurrentStep.AddChildElementbyname "EXCEL", sExcelName
	
	REM oTCResultdoc.SaveFile  sResultFileName

	
REM '	If Err.Number <> 0 Then
REM '		iModuleFail = 1
REM '		AddVerificationInfoToResult "Error",Err.Description
REM '	End If
REM End Function

'#################################################################################################
'         Function Name : UpdateResultSummaryTotal
'         Description :  This function updates Total Test case number in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function UpdateResultSummaryTotal(TOTAL)

	Set SummaryElement = oResultDoc.ChildElementsByPath("/RESULT/SUMMARY")
	set CurrentSummaryElement=SummaryElement.ItemByName("SUMMARY",SummaryElement.Count)
	CurrentSummaryElement.ChildElements().ItemByName("TOTAL",1).SetValue(TOTAL)

	oResultdoc.SaveFile  sResultFileName
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function
'#################################################################################################
'         Function Name : UpdateResultSummaryTotalPass
'         Description : This function updates total test case pass number in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function UpdateResultSummaryTotalPass()
	currVal=0
	Set SummaryElement = oResultDoc.ChildElementsByPath("/RESULT/SUMMARY")
	set CurrentSummaryElement=SummaryElement.ItemByName("SUMMARY",SummaryElement.Count)
	currVal = Cint(CurrentSummaryElement.ChildElements().ItemByName("TOTALPASS",1).Value())
	CurrentSummaryElement.ChildElements().ItemByName("TOTALPASS",1).SetValue(Cstr(currVal+1))

   ' oResultdoc.SaveFile  sResultFileName
'   If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function
'#################################################################################################
'         Function Name : UpdateResultSummaryTotalFail
'         Description :  This function updates total test case fail number in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function UpdateResultSummaryTotalFail()
	currVal=0
	Set SummaryElement = oResultDoc.ChildElementsByPath("/RESULT/SUMMARY")
	set CurrentSummaryElement=SummaryElement.ItemByName("SUMMARY",SummaryElement.Count)
	currVal = Cint(CurrentSummaryElement.ChildElements().ItemByName("TOTALFAIL",1).Value())
	CurrentSummaryElement.ChildElements().ItemByName("TOTALFAIL",1).SetValue(Cstr(currVal+1))

	'oResultdoc.SaveFile  sResultFileName
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function
'#################################################################################################
'         Function Name : AddTestCaseToResult
'         Description : This function adds a test case in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function AddTestCaseToResult(TCID,TCNAME,TCRESULT)
   oRoot.AddChildElementbyname "TESTCASE",""

	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
	lastTCNo = TestCases.Count 

	set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)
	CurrentTC.AddChildElementbyname "TCID",TCID
	CurrentTC.AddChildElementbyname "NAME",TCNAME
	CurrentTC.AddChildElementbyname "RESULT",TCRESULT

	oResultdoc.SaveFile  sResultFileName
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function

'#################################################################################################
'         Function Name : AddTestCaseToTCResultFile
'         Description : This function adds a test case in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
REM Public Function AddTestCaseToTCResultFile(TCID,TCNAME,TCRESULT)
   REM oTCRoot.AddChildElementbyname "TESTCASE",""

	REM Set TestCases = oTCResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
	REM lastTCNo = TestCases.Count 

	REM set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)
	REM CurrentTC.AddChildElementbyname "TCID",TCID
	REM CurrentTC.AddChildElementbyname "NAME",TCNAME
	REM CurrentTC.AddChildElementbyname "RESULT",TCRESULT

	REM oTCResultdoc.SaveFile  sResultFileName
REM '	If Err.Number <> 0 Then
REM '		iModuleFail = 1
REM '		AddVerificationInfoToResult "Error",Err.Description
REM '	End If
REM End Function

'#################################################################################################
'         Function Name : AddStepToResult
'         Description : This function adds test case step in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function AddStepToResultOld(STEPNO,STEPNAME,STEPRESULT)
	
	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
	lastTCNo = TestCases.Count 

	Set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)
	CurrentTC.AddChildElementbyname "STEPS",""

	Set Steps = CurrentTC.ChildElements()
	lastStep = Steps.Count 
	Set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)

	CurrentStep.AddChildElementbyname "STEPNO",STEPNO
	CurrentStep.AddChildElementbyname "NAME",STEPNAME
	CurrentStep.AddChildElementbyname "RESULT",STEPRESULT
	'sExcelName = "C:\QA_Auto_QTP\Campaign\Data\" & Environment("sDataFileName") & "#" & Environment("sDataSheetName") & "!a1"
	'CurrentStep.AddChildElementbyname "EXCEL","TEST"

	oResultdoc.SaveFile  sResultFileName
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function

Function AddStepToResult(STEPNO,STEPNAME,STEPRESULT)

	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE/TCID[text()='"+sTCID+"']/..")
	cnt = TestCases.Count 

	Dim CurrentTC
'	For i=1 to cnt
	Set CurrentTC=TestCases.Item(1)

'	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
'	cnt = TestCases.Count 
'	sTCID = adoRS("TCID") & ""
'	Dim CurrentTC
'	For i=1 to cnt
'		Set CurrentTC=TestCases.ItemByName("TESTCASE",i)
'		If  CurrentTC.getValueByXpath("./TCID")= sTCID then
		CurrentTC.AddChildElementbyname "STEPS",""
		Set Steps = CurrentTC.ChildElements()
		lastStep = Steps.Count 
		Set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)
	
		CurrentStep.AddChildElementbyname "STEPNO",STEPNO
		CurrentStep.AddChildElementbyname "NAME",STEPNAME
		CurrentStep.AddChildElementbyname "RESULT",STEPRESULT
		oResultdoc.SaveFile  sResultFileName
'			Exit For
'		End If
'	
'	Next
End Function



'#################################################################################################
'         Function Name : AddStepToResult
'         Description : This function adds test case step in the xml result file
'         Created By : 
'         Creation Date :        
''##################################################################################################
'Public Function AddStepToResultPostSubmission(STEPNO,STEPNAME,STEPRESULT)
'	
'	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
'	lastTCNo = TestCases.Count 
'
'	Set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)
'	CurrentTC.AddChildElementbyname "STEPS",""
'
'	Set Steps = CurrentTC.ChildElements()
'	lastStep = Steps.Count 
'	Set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)
'
'	CurrentStep.AddChildElementbyname "STEPNO",STEPNO
'	CurrentStep.AddChildElementbyname "NAME",STEPNAME
'	CurrentStep.AddChildElementbyname "RESULT",STEPRESULT
'	'sExcelName = "C:\QA_Auto_QTP\Campaign\Data\" & Environment("sDataFileName") & "#" & Environment("sDataSheetName") & "!a1"
'	'CurrentStep.AddChildElementbyname "EXCEL","TEST"
'
'	oResultdoc.SaveFile  sResultFileName
''	If Err.Number <> 0 Then
''		iModuleFail = 1
''		AddVerificationInfoToResult "Error",Err.Description
''	End If
'End Function

'#################################################################################################
'         Function Name : AddStepToResult
'         Description : This function adds test case step in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
REM Public Function AddStepToTCResultFile(STEPNO,STEPNAME,STEPRESULT)
	
	REM Set TestCases = oTCResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
	REM lastTCNo = TestCases.Count 

	REM Set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)
	REM CurrentTC.AddChildElementbyname "STEPS",""

	REM Set Steps = CurrentTC.ChildElements()
	REM lastStep = Steps.Count 
	REM Set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)

	REM CurrentStep.AddChildElementbyname "STEPNO",STEPNO
	REM CurrentStep.AddChildElementbyname "NAME",STEPNAME
	REM CurrentStep.AddChildElementbyname "RESULT",STEPRESULT
	REM 'sExcelName = "C:\QA_Auto_QTP\Campaign\Data\" & Environment("sDataFileName") & "#" & Environment("sDataSheetName") & "!a1"
	REM 'CurrentStep.AddChildElementbyname "EXCEL","TEST"

	REM oTCResultdoc.SaveFile  sResultFileName
REM '	If Err.Number <> 0 Then
REM '		iModuleFail = 1
REM '		AddVerificationInfoToResult "Error",Err.Description
REM '	End If
REM End Function


'#################################################################################################
'         Function Name : AddFailureDescToResult
'         Description :  This function adds failure description to a step in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function AddFailureDescToResult(sDESC,sOBJECT, sNAME,sMETHOD,sARGUMENTS)
	
'	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
'	lastTCNo = TestCases.Count 
'
'	Set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)
'	'CurrentTC.AddChildElementbyname "STEPS",""
'
'	Set Steps = CurrentTC.ChildElements()
    Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE/TCID[text()='"+sTCID+"']/..")
	cnt = TestCases.Count 

	Dim CurrentTC
'	For i=1 to cnt
	Set CurrentTC=TestCases.Item(1)
	Set Steps = CurrentTC.ChildElements()
	lastStep = Steps.Count 
	Set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)
	
	CurrentStep.AddChildElementbyname "DESC",sDESC
	CurrentStep.AddChildElementbyname "OBJECT",sOBJECT
	CurrentStep.AddChildElementbyname "OBJNAME",sNAME
	CurrentStep.AddChildElementbyname "METHOD",sMETHOD
	CurrentStep.AddChildElementbyname "ARGUMENTS",sARGUMENTS

	Dim sErrImagePath
	Dim sErrImgName
	sErrImgName = TimeStamp_fn() & "TCID" & sTCID & ".png"
	'By Ankit for change in reporting
'	sErrImagePath = sProjectDir & "Driver\Results\" & sProduct_G  & "\"   & sErrImgName
	sErrImagePath = Environment("RunFolderPath") & "\" & Environment("ResFldrName") & "\"   & sErrImgName
	sErrImagePath = Environment("RunFolderPath") & "\"   & sErrImgName
	'sExcelName = "C:\QA_Auto_QTP\Campaign\Data\" & Environment("sDataFileName") & "#" & Environment("sDataSheetName") & "!a1"
	'sExcelName = "C:\QA_Auto_QTP\Campaign\Data\openfile.vbs " & Environment("sDataFileName") & " " & Environment("sDataSheetName")

	Desktop.CaptureBitmap sErrImagePath,1
	CurrentStep.AddChildElementbyname "IMAGE",sErrImgName
	'CurrentStep.AddChildElementbyname "EXCEL",sExcelName		'Added by Anunay
	
	oResultdoc.SaveFile  sResultFileName
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function


'#################################################################################################
'         Function Name : AddFailureDescToTCResultFile
'         Description :  This function adds failure description to a step in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function AddFailureDescToTCResultFile(sDESC,sOBJECT, sNAME,sMETHOD,sARGUMENTS)
	
'	Set TestCases = oTCResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
'	lastTCNo = TestCases.Count 
'
'	Set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)

	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE/TCID[text()='"+sTCID+"']/..")
	cnt = TestCases.Count 

	Dim CurrentTC
'	For i=1 to cnt
	Set CurrentTC=TestCases.Item(1)
	'CurrentTC.AddChildElementbyname "STEPS",""

	Set Steps = CurrentTC.ChildElements()
	lastStep = Steps.Count 
	Set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)
	
	CurrentStep.AddChildElementbyname "DESC",sDESC
	CurrentStep.AddChildElementbyname "OBJECT",sOBJECT
	CurrentStep.AddChildElementbyname "OBJNAME",sNAME
	CurrentStep.AddChildElementbyname "METHOD",sMETHOD
	CurrentStep.AddChildElementbyname "ARGUMENTS",sARGUMENTS

	Dim sErrImagePath
	Dim sErrImgName
	sErrImgName = TimeStamp_fn() & "TCID" & sTCID & ".png"
	'By Ankit for change in reporting
'	sErrImagePath = sProjectDir & "Driver\Results\" & sProduct_G  & "\"   & sErrImgName
	sErrImagePath = Environment("RunFolderPath") & "\" & Environment("ResFldrName") & "\"   & sErrImgName
	'sExcelName = "C:\QA_Auto_QTP\Campaign\Data\" & Environment("sDataFileName") & "#" & Environment("sDataSheetName") & "!a1"
	'sExcelName = "C:\QA_Auto_QTP\Campaign\Data\openfile.vbs " & Environment("sDataFileName") & " " & Environment("sDataSheetName")

	Desktop.CaptureBitmap sErrImagePath,1
	CurrentStep.AddChildElementbyname "IMAGE",sErrImgName
	'CurrentStep.AddChildElementbyname "EXCEL",sExcelName		'Added by Anunay
	
	oTCResultDoc.SaveFile  sResultFileName
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function

'#################################################################################################
'         Function Name : AddVerificationInfoToResult
'         Description : This function adds verification(execution flow) information in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function AddVerificationInfoToResult(DESC,RESULT)
'	
'	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
'	lastTCNo = TestCases.Count 
'
'	Set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)

	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE/TCID[text()='"+sTCID+"']/..")
	cnt = TestCases.Count 

	Dim CurrentTC
'	For i=1 to cnt
	Set CurrentTC=TestCases.Item(1)

	Set Steps = CurrentTC.ChildElements()
	lastStep = Steps.Count 
	Set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)

	CurrentStep.AddChildElementbyname "VERIFICATION",""

	Set VerificationElements = CurrentStep.ChildElements()
	'lastData = VerificationElements.Count 
	lastData=CurrentStep.GetNumDescendantElemByName("VERIFICATION")
	Set CurrentData=VerificationElements.ItemByName("VERIFICATION",lastData)
	
	CurrentData.AddChildElementbyname "NAME",DESC
	CurrentData.AddChildElementbyname "VALUE",RESULT
	

	oResultdoc.SaveFile  sResultFileName
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'	End If
End Function


'#################################################################################################
'         Function Name : AddVerificationInfoToResult
'         Description : This function adds verification(execution flow) information in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
REM Public Function AddVerificationInfoToTCResultFile(DESC,RESULT)
	
	REM Set TestCases = oTCResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
	REM lastTCNo = TestCases.Count 

	REM Set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)

	REM Set Steps = CurrentTC.ChildElements()
	REM lastStep = Steps.Count 
	REM Set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)

	REM CurrentStep.AddChildElementbyname "VERIFICATION",""

	REM Set VerificationElements = CurrentStep.ChildElements()
	REM 'lastData = VerificationElements.Count 
	REM lastData=CurrentStep.GetNumDescendantElemByName("VERIFICATION")
	REM Set CurrentData=VerificationElements.ItemByName("VERIFICATION",lastData)
	
	REM CurrentData.AddChildElementbyname "NAME",DESC
	REM CurrentData.AddChildElementbyname "VALUE",RESULT
	

	REM oTCResultDoc.SaveFile  sResultFileName
REM '	If Err.Number <> 0 Then
REM '		iModuleFail = 1
REM '	End If
REM End Function

'#################################################################################################
'         Function Name : AddStepDataToResult
'         Description : This function adds test data to a step in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function AddStepDataToResult(FIELD,FIELDVALUE)
	
'	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
'	lastTCNo = TestCases.Count 
'
'	Set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)

    Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE/TCID[text()='"+sTCID+"']/..")
	cnt = TestCases.Count 

	Dim CurrentTC
'	For i=1 to cnt
	Set CurrentTC=TestCases.Item(1)

	Set Steps = CurrentTC.ChildElements()
	lastStep = Steps.Count 
	Set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)

	CurrentStep.AddChildElementbyname "DATA",""

	Set DataElements = CurrentStep.ChildElements()
	lastData = DataElements.Count 
	Set CurrentData=DataElements.ItemByName("DATA",lastData-3)
	
	CurrentData.AddChildElementbyname "NAME",FIELD
	CurrentData.AddChildElementbyname "VALUE",FIELDVALUE
	

	oResultdoc.SaveFile  sResultFileName
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function
'#################################################################################################
'         Function Name : GetDataUsed
'         Description : This function retrives data that is used in a test step
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function GetDataUsed()
	Dim strSQL
	Dim adoFunctionDesc
	err.clear	
	strSQL = "Select DataFileName,DataSheetName from [FunctionDescription$] where FunctionName = '" & sCurrent_Module & "'"
	Set adoFunctionDesc =ExecExcQry(sProductDir & "FunctionsCreated.xls", strSQL)
	If adoFunctionDesc.Eof = "False" then
		sDataFileName =adoFunctionDesc("DataFileName")
		Environment("sDataFileName") = sDataFileName
		sDataSheetName =adoFunctionDesc("DataSheetName")
		Environment("sDataSheetName") = sDataSheetName
	End if	
	adoFunctionDesc.Close	
	Set adoFunctionDesc = Nothing
	
	If trim(sDataFileName) <> "" AND trim(sDataSheetName) <> "" then
		Dim adoFunctionData
		strSQL = "Select * from [" & sDataSheetName & "$] where [RowNo]=" & iRowNo
		Set adoFunctionData =ExecExcQry(sProductDir & "Data\" & sDataFileName , strSQL)
	
		If adoFunctionData.Eof = "False" then
			Do while not adoFunctionData.Eof
				For i=0 to adoFunctionData.fields.Count-1
			   ' If GetDataUsed = "" Then
					call AddStepDataToResult(adoFunctionData(i).Name ,ReplaceInData(cStr(adoFunctionData(i) & "")))
				'Else	
				   ' GetDataUsed =  sFunctionData & " , " & adoFunctionData(i).Name & " = " & adoFunctionData(i)
				'End If
				Next
				adoFunctionData.MoveNext
			Loop
		End if
		adoFunctionData.Close	
		call AddStepDataToResult("Debug Info", "Data File: " & sDataFileName & " | Data Sheet: " & sDataSheetName)
		Set adoFunctionData = Nothing
	End if
	'strSQL = "Update [RESULT$] SET DATA = '" & sFunctionData & "' WHERE TCID = '" & sTCID & "' AND Name = '" & sCurrent_Module & "'"
	'Set adoResult =ExecExcQry(sProjectDir & "Results\" & FunctionsCreated.xls , strSQL)
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function
'#################################################################################################
'         Function Name : WriteResultSummary
'         Description : This function writes result summary to a flat file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function WriteResultSummary()
	'delete file that exists (of the previous run)
	Dim objFile, objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(sProjectDir & "output.txt")
	objFile.Delete ()

	'Create and write to the Output File
	DIM fso, OutputFile
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set OutputFile = fso.CreateTextFile(sProjectDir & "output.txt", True)
	OutputFile.WriteLine(iAtLeastOneTestFailed & " number of test cases failed") 
	OutputFile.WriteLine(iNumberOfTestCasesExecuted  & " were the total number of Executed Test Cases") 
	OutputFile.Close
 

	'delete file that exists (of the previous run)
'	Dim objExcelFile,objExl
'	Set objExl = CreateObject("Excel.Application")
'	Set objExcelFile = objFSO.GetFile(sProjectDir & "AutoResults.xls")
'	objExcelFile.Delete ()
	
	 
	'to copy the Result file in the C:\QA_Auto_QTP directory
'	Set sObjExcel = CreateObject("Excel.Application")
'	sResultExcel = Environment.Value("SavedResultPath") 
'	Set sObjWbk = sObjExcel.Workbooks.open (sResultExcel)
'	sObjWbk.SaveAs  sProjectDir & "AutoResults.xls"
'	sObjWbk.Save
'	sObjExcel.Quit      
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function


'#################################################################################################
'         Function Name : CaptureScreenshot
'         Description :  This function adds failure description to a step in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function CaptureScreenshot
	
	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
	lastTCNo = TestCases.Count 

	Set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)
	'CurrentTC.AddChildElementbyname "STEPS",""

	Set Steps = CurrentTC.ChildElements()
	lastStep = Steps.Count 
	Set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)
	

	Dim sErrImagePath
	Dim sErrImgName
	sErrImgName = TimeStamp_fn() & "TCID" & sTCID & ".png"
	'By Ankit for change in reporting
'	sErrImagePath = sProjectDir & "Driver\Results\" & sProduct_G  & "\"   & sErrImgName
	sErrImagePath = Environment("RunFolderPath") & "\" & Environment("ResFldrName") & "\"   & sErrImgName

	'sExcelName = "C:\QA_Auto_QTP\Campaign\Data\" & Environment("sDataFileName") & "#" & Environment("sDataSheetName") & "!a1"
	'sExcelName = "C:\QA_Auto_QTP\Campaign\Data\openfile.vbs " & Environment("sDataFileName") & " " & Environment("sDataSheetName")

	Desktop.CaptureBitmap sErrImagePath,1
	CurrentStep.AddChildElementbyname "IMAGE",sErrImgName
	'CurrentStep.AddChildElementbyname "EXCEL",sExcelName		'Added by Anunay
	
	oResultdoc.SaveFile  sResultFileName
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function

'#################################################################################################
'         Function Name : RemainingTestCaseSteps
'         Description :  This function adds failure description to a step in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function RemainingTestCaseSteps
	
	'Added to display remaining steps==========================
	Do While sCurrent_Value <> ""
		Err.Clear
		iCol = iCol + 1
		sModule = "module" & icol
		sCurrent_Value = adoRS(sModule)
		If sCurrent_Value <> "" Then
			sMyArray = split(sCurrent_Value,";",-1,1)
			sCurrent_Module = sMyArray(0)
			If Ubound(sMyArray) <> 0 Then
				iRowNo = sMyArray(1)
			End If

			Call AddStepToResult(cstr(iCol),sCurrent_Module,"Not Executed")

			'Print "Error : " & Err.Description

			If Ubound(sMyArray) <> 0 Then
				call GetDataUsed()
			End If
			Call UpdateStepStatusInResult("Not Executed")
			'Print "Error : " & Err.Description
			'Print sCurrent_Module
		End If
	Loop
End Function

'#################################################################################################
'         Function Name : AddPassDescToResult
'         Description :  This function adds failure description to a step in the xml result file
'         Created By : 
'         Creation Date :        
REM '##################################################################################################
REM Public Function AddPassDescToResult
	
	REM Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
	REM lastTCNo = TestCases.Count 

	REM Set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)
	REM 'CurrentTC.AddChildElementbyname "STEPS",""

	REM Set Steps = CurrentTC.ChildElements()
	REM lastStep = Steps.Count 
	REM Set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)
	


	REM Dim sErrImagePath
	REM Dim sErrImgName
	REM sErrImgName = TimeStamp_fn() & "TCID" & sTCID & ".png"
	REM 'By Ankit for change in reporting
REM '	sErrImagePath = sProjectDir & "Driver\Results\" & sProduct_G  & "\"   & sErrImgName
	REM sErrImagePath = TCFolderPath  & "\"   & sErrImgName

	REM 'sExcelName = "C:\QA_Auto_QTP\Campaign\Data\" & Environment("sDataFileName") & "#" & Environment("sDataSheetName") & "!a1"
	REM 'sExcelName = "C:\QA_Auto_QTP\Campaign\Data\openfile.vbs " & Environment("sDataFileName") & " " & Environment("sDataSheetName")

	REM Desktop.CaptureBitmap sErrImagePath,1
	REM CurrentStep.AddChildElementbyname "IMAGE",sErrImagePath
	REM 'CurrentStep.AddChildElementbyname "EXCEL",sExcelName		'Added by Anunay
	
	REM oResultdoc.SaveFile  sResultFileName
REM '	If Err.Number <> 0 Then
REM '		iModuleFail = 1
REM '		AddVerificationInfoToResult "Error",Err.Description
REM '	End If
REM End Function

'#################################################################################################
'         Function Name : AddPassDescToResult
'         Description :  This function adds failure description to a step in the xml result file
'         Created By : 
'         Creation Date :        
'##################################################################################################
Public Function AddPassDescToResult
	
'	Set TestCases = oTCResultDoc.ChildElementsByPath("/RESULT/TESTCASE")
'	lastTCNo = TestCases.Count 
'
'	Set CurrentTC=TestCases.ItemByName("TESTCASE",lastTCNo)
	'CurrentTC.AddChildElementbyname "STEPS",""

	Set TestCases = oResultDoc.ChildElementsByPath("/RESULT/TESTCASE/TCID[text()='"+sTCID+"']/..")
	cnt = TestCases.Count 

	Dim CurrentTC
'	For i=1 to cnt
	Set CurrentTC=TestCases.Item(1)

	Set Steps = CurrentTC.ChildElements()
	lastStep = Steps.Count 
	Set CurrentStep=Steps.ItemByName("STEPS",lastStep-3)
	


	Dim sErrImagePath
	Dim sErrImgName
	sErrImgName = TimeStamp_fn() & "TCID" & sTCID & ".png"
	'By Ankit for change in reporting
'	sErrImagePath = sProjectDir & "Driver\Results\" & sProduct_G  & "\"   & sErrImgName
'	sErrImagePath = TCFolderPath  & "\"   & sErrImgName
	sErrImagePath = Environment("RunFolderPath") & "\" & Environment("ResFldrName") & "\"   & sErrImgName

	'sExcelName = "C:\QA_Auto_QTP\Campaign\Data\" & Environment("sDataFileName") & "#" & Environment("sDataSheetName") & "!a1"
	'sExcelName = "C:\QA_Auto_QTP\Campaign\Data\openfile.vbs " & Environment("sDataFileName") & " " & Environment("sDataSheetName")

	Desktop.CaptureBitmap sErrImagePath,1
	CurrentStep.AddChildElementbyname "IMAGE",sErrImgName
	'CurrentStep.AddChildElementbyname "EXCEL",sExcelName		'Added by Anunay
	
	oResultDoc.SaveFile  sResultFileName
'	If Err.Number <> 0 Then
'		iModuleFail = 1
'		AddVerificationInfoToResult "Error",Err.Description
'	End If
End Function


Public Function CreateTCFolder(sTCName)
   Dim str
   Const LETTERS = "abcdefghijklmnopqrstuvwxyz0123456789" 
   For i=1 to 10
	   str=str & mid (LETTERS,RandomNumber(1,Len(LETTERS)),1)

   Next
   Set fso1= CreateObject("Scripting.FileSystemObject")
   tcFolderPath=Environment.Value("RunFolderPath") & "\" & sTCName & "_" & str
   Environment.Value("ResFldrName")=sTCName & "_" & str

   Set f = fso1.CreateFolder(tcFolderPath)
   TCFolderPath = f.Path

'   resultPath=sProjectDir & "Driver\Results\" & sProduct_G  & "\"
'	fso1.CopyFile resultPath & "minus.gif", TCFolderPath & "\", True
'	fso1.CopyFile resultPath & "plus.gif", TCFolderPath & "\", True
'	fso1.CopyFile resultPath & "leaf.gif", TCFolderPath & "\", True
'	fso1.CopyFile resultPath & "Result.dtd", TCFolderPath & "\", True
'	fso1.CopyFile resultPath & "Result.xsl", TCFolderPath & "\", True

End Function




'#################################################################################################
' 	Function Name : WriteToOutputXcelFile_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function WriteToOutputXcelFile_fn()
'   	Set adoExcConn = CreateObject("ADODB.Connection")
	Err.Clear
'    strSQL = "SELECT max([RowNo]) as RowNo FROM [Sheet1$]"
'	strSQL = "SELECT top 1 [RowNo] as RowNo FROM [Sheet1$] order by [RowNo] desc"

'	Set adoData = ExecExcQry(Environment("RunFolderPath")&"\"&Environment("OutputFile"),strSQL)

'	For each fld in adoData.Fields
'		msgbox fld.name
'
'	Next
'	DictionaryTest_G.Item("AccountNo")="abk"
'	DictionaryTest_G.Item("OrderNo")="giuksi"
'	DictionaryTest_G.Item("Status")="od"
		On error resume next
		row=outputRowNum
		If isNull(row) Then
			row=0
		End If
		sResFldrName=Environment("ResFldrName")
		If isEmpty(sResFldrName) Then
			sResFldrName=""
		End If
		sModule=Environment("Module")
		sAccountNo=""
		If DictionaryTest_G.Exists("AccountNo") Then
			sAccountNo=DictionaryTest_G("AccountNo")
		End If
		sOrderNo=""
		If DictionaryTest_G.Exists("OrderNo") Then
			sOrderNo=DictionaryTest_G.Item("OrderNo")
		End If
		sMSISDN=""
		If DictionaryTest_G.Exists("ACCNTMSISDN") Then
			sMSISDN=DictionaryTest_G.Item("ACCNTMSISDN")
		End If
		sStatus=""
		If DictionaryTest_G.Exists("Status") Then
			sStatus=DictionaryTest_G.Item("Status")
		End If
		err.clear
'		strSQL="insert into [Sheet1$] ([RowNo],[Env],[Module],[TCName],[AccountNo],[OrderNo],[MSISDN],[Status],[TCResult],[ResFldrName]) values(5,'a','a','a','a','a','a','a','a','a')"
		strSQL="insert into [Sheet1$] ([RowNo],[Env],[Module],[TCID],[TCName],[AccountNo],[OrderNo],[MSISDN],[Status],[TCResult],[ResFldrName]) values("&Cint(row)+1&",'"&sEnv&"','"&sModule&"','"&sTCID&"','"&sTestCaseName&"','"&sAccountNo&"','"&sOrderNo&"','"&sMSISDN&"','"&sStatus&"','"&sTCResult&"','"&sResFldrName&"')"
'		strSQL="insert into [Sheet1$] ([RowNo],[Env],[Module],[TCName],[AccountNo],[OrderNo],[MSISDN],[Status],[TCResult],[ResFldrName]) values(1,'a','a','a','','a','a','a','a','a')"
		UpdateExcQry Environment("RunFolderPath")&"\"&Environment("OutputFile"),strSQL

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function



'#################################################################################################
' 	Function Name : UpdateOutputXcelFile_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function UpdateOutputXcelFile_fn()
		Err.Clear
		sStatus=""
		If DictionaryTest_G.Exists("Status") Then
			sStatus=DictionaryTest_G.Item("Status")
		End If


		If Ucase(sTCResult)<>"PASS" and sTCResult<>"Not Executed" Then
			sTCResult = "Fail"
		elseif sTCResult = "Not Executed" Then
			sTCResult = "Not Executed"
		End If

		strSQL="update [Sheet1$] set [Status]='" & sStatus & "',[TCResult]='" & sTCResult & "' where [RowNo]='" &  outputRowNum & "'"

		UpdateExcQry Environment("RunFolderPath")&"\"&Environment("OutputFile"),strSQL

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If


End Function

'#################################################################################################
' 	Function Name : RetrieveOrderDetailsFromOutputXcelFile_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function RetrieveOrderDetailsFromOutputXcelFile_fn()
'  	Set adoExcConn = CreateObject("ADODB.Connection")
	Err.Clear
    strSQL = "SELECT * FROM [Sheet1$] WHERE  [TCName]='" & sTestCaseName & "' and [TCID]= '" & sTCID &"'"

	Set adoData = ExecExcQry(Environment("RunFolderPath")&"\"&Environment("OutputFile"),strSQL)
		outputRowNum = adoData("RowNo") & ""

		Environment("ResFldrName")=adoData("ResFldrName") & ""

		Environment("Module")=adoData("Module") & ""
'		sAccountNo=""
		If DictionaryTest_G.Exists("AccountNo") Then
			DictionaryTest_G.Item("AccountNo")=adoData("AccountNo") & ""
		else
			DictionaryTest_G.add "AccountNo",adoData("AccountNo") & ""
		End If
'		sOrderNo=""
		If DictionaryTest_G.Exists("OrderNo") Then
			DictionaryTest_G.Item("OrderNo")=adoData("OrderNo") & ""
		else
			DictionaryTest_G.add "OrderNo",adoData("OrderNo") & ""
		End If
'		sMSISDN=""
		If DictionaryTest_G.Exists("ACCNTMSISDN") Then
			DictionaryTest_G.Item("ACCNTMSISDN")=adoData("MSISDN") & ""
		else
			DictionaryTest_G.add "ACCNTMSISDN",adoData("MSISDN") & ""
		End If
'		sStatus=""
		If DictionaryTest_G.Exists("Status") Then
			DictionaryTest_G.Item("Status")=adoData("Status") & ""
		else
			DictionaryTest_G.add "Status",adoData("Status") & ""
		End If
		sTCResult = adoData("TCResult") & ""
		sEnv=adoData("Env") & ""
End Function


Public Function CaptureSnapshot
	sImgName = TimeStamp_fn() & "TCID" & sTCID & ".png"
'	sImagePath = sProjectDir & "Driver\Results\" & sProduct_G  & "\"   & sImgName
	sImagePath = Environment("RunFolderPath") & "\" & Environment("ResFldrName") & "\"   & sImgName

	Desktop.CaptureBitmap sImagePath,1
End Function
