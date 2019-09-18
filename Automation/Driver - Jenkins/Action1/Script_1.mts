'*********************************************************************************************
'File Name: Driver SetSecurityPolicy_fn SetSecurityPolicy_fn CreateUser_fn AssignRoleToUser_fn FillDataInDimensionTabOfEPB_fn DeviceReplayfunction FillDataInSouceTabEPB_fn;11;0 SelectCFToExtract_fn;16;0 CheckToolTip_fn;315;0 FillDataInDimensionTabOfEPB_fn;3;0
'FillDataInEditSegInCSPB_fn normalclick FillCampaignInfo_fn FillCampaignBasicInfo_fn CreateProject_fn;15308;0 SearchClickOnProject_fn;12110;0 CreateProject_fn;15308;0 SearchClickOnProject_fn;12150; VerifyObjectExist_fn;3013;0 SelectDimensionTableKeyField_fn;7;0
'SelectDimensionTableKeyField_fn ViewLog_fn;31; SetSecurityPolicy_fn ;609;0 ClickOnWebTblCell_fn;4;0 ActionOnMonitoring_fn;3;0 CheckMonitoring_fn;2;0 MatchLinkTargetCells_fn;2;0 FillDataInTargetCell_fn;350;0 FillDataInSouceTabOfSelectPB_fn;351;0
'FillDataInProfileOption_fn VerifyDataInProfileSelectedField_fn EnterDataInRowForGivenColumn_fn FillDataInCustomAttribute_fn;153; LocalListClick FillDataInTargetCell_fn;356;0 CheckRecordPresent_fn VerifyObjectExist_fn;50121;0 Fil
'lDataInAdvanceSettingForSelectPB_fn;256 FillDataInSouceTabOfSelectPB_fn;351;0 PointAndClick_fn;7 ClickOnAProject_fn;12013;0 SetObjRepository EnterDataInRowForGivenColumn_fn






'DialogErrMsg_fn;8;0 PublishCopyForm_fn DragAttributeToGrid_fn;5;0


'Browser("Affinium Plan").Page("Affinium Plan").Frame("content").Link("Shared Attributes").Click
'
'Browser("Unica Distributed Marketing").Page("Unica Distributed Marketing").Frame("content").WebElement("html tag:=TR","innertext:=Grid03.*").Link("html tag:=A","text:=Enable").Click

'Action Name: TestSetExec   
'Author: Pramod Moorjani / Ganesh Pawar FillDataInSession_fn InsertProcessBox_fn SelectNewTableType_fn
'Start Date: Nov 2006
'Brief description: This Action is used to execute the list of Test Cases
'*********************************************************************************************  VerifyListData_fn FillDataInSourceTabRPB_fn FillDataInTrackPBSource_fn SpecifySourceFile_fn DialogErrMsg_fn FillDataSchedulePB_fn CheckMonitoring_fn;38;0

Call ClearErrorInVbs_fn()
Call  InitializeVariables()
Call setpath(Environment.Value("Product")) 

Dim xmlReportPath, iTestCnt
Dim sPrefixXmlResultFileWith

sPrefixXmlResultFileWith = Environment.Value("ResultFileName")

'For Reporting results to xml
'*********************************************************************************************
xmlReportPath= CreateAXMLResultfile(sPrefixXmlResultFileWith)

'Query for executing test cases
'*********************************************************************************************
sTestCaseList = Environment.Value("TestCaseFile") 
strSQL =Environment.Value("Query")
Set adoRS = ExecExcQry(sTestCaseList , strSQL)

If  adoRS.Eof  Then ' No test case seleted for execution
	ExitAction
End If

Call UpdateResultSummaryTotal(Cstr(adoRS.RecordCount))

iTestCnt = 0		'Count of executed testcases in a module
adoRS.MoveFirst

Do While Not adoRS.EOF
	iTestCnt = iTestCnt + 1
	Environment("iTestCnt") = iTestCnt
	sTCResult = 0
	sTestCaseName=""
	Call ClearErrorInVbs_fn()
   
	sTCResult = RunAction ("ExecuteTestCases", oneIteration)

	If sTCResult = "Pass"   Then
		Call UpdateTestCaseStatusInResult("PASS")
		Reporter.ReportEvent micPass, sTestCaseName, "All modules of the" & sTestCaseName & " passed "
		iContinuousFailNumber = 0
	Else
		If  sTCResult<>"Not Executed" Then
			Call UpdateTestCaseStatusInResult("FAIL") 
			IsPreTCFailed ="Yes"
			Dictionary_G.Add  sTestCaseName ,"FAIL"
			iContinuousFailNumber = iContinuousFailNumber +1
		End If
	End if
   
	iNumberOfTestCasesExecuted = iNumberOfTestCasesExecuted +  1
	If sTCResult <> "Pass" Then
		iAtLeastOneTestFailed =  iAtLeastOneTestFailed + 1
		If iContinuousFailNumber > 3 Then
			SendAlerts()
			iContinuousFailNumber = 0
		Exit Do
		End If
	End if
	adoRS.MoveNext
Loop

adoRS.Close
Set adoRS = Nothing

If  iAtLeastOneTestFailed <> 0  Then
	Reporter.ReportEvent 1, "Total failures", "The number of test cases that failed were ' " & iAtLeastOneTestFailed  & " '"   
End If
Call WriteResultSummary()








































































































































