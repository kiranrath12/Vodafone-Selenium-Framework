'#################################################################################################
' 	Function Name : FileTransferToLocal_fn
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function FileTransferToLocal_fn()

    strSQL = "SELECT * FROM [FileTransferToLocal$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\PSCP.xls",strSQL)
	call SetObjRepository ("Putty",sProductDir & "Resources\")
	
	'sOR

	cstrSftp = """C:\PuTTY\pscp.exe"""
	sHostName = adoData("Host"&sEnv) & ""
	sRemotePath = adoData("RemotePath") & ""

'	DictionaryTest_G("ReportFile") = "brm_po_voice_002203_1_20170307133805.dat"
'	DictionaryTest_G("ReportFile") = "Proof_and_balance_report_custom_201611_20161205144909_01.xls"
	sRemotePathFile = sRemotePath & "/" & DictionaryTest_G("ReportFile")

	sReportPath = Environment("RunFolderPath") & "\" & Environment("ResFldrName")
'sLogPath = Environment("RunFolderPath")
'	sReportPath ="C:\Ankit"

    strCommand = cstrSftp & " -sftp -l " & sBRMPuttyLogin & " -pw " & sBRMPuttyPassphrase & " " & sHostName & ":" & sRemotePathFile & " " & sReportPath
	Set objShell = CreateObject("WScript.Shell")
	 objShell.Run strCommand, 1
'	 objShell.Run strCommand, 1,True

	wait 2
	If Window("PSCP").exist(0) Then
		Window("PSCP").Activate
		Window("PSCP").Type "y"
		Window("PSCP").Type micReturn
	End If
	
	wait 2
	
	Set FSO = CreateObject("Scripting.FileSystemObject") 

	sFileName = sReportPath & "\" & DictionaryTest_G("ReportFile")
	
	boolRC = FSO.FileExists(sFileName) 
	
	If Not boolRC Then 
		AddVerificationInfoToResult  "Error" ,"File:-" & DictionaryTest_G("ReportFile") & " is not downloaded successfully."
		iModuleFail = 1
	End If 

	Set FSO = Nothing 'release an object 	

	CaptureSnapshot()

End Function


'#################################################################################################
' 	Function Name : CompareXcelReports_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function CompareXcelReports_fn()

    strSQL = "SELECT * FROM [CompareXcelReports$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\PSCP.xls",strSQL)
	
	'sOR
	sCell = adoData( "Cell")  & ""
	sResult = adoData( "Result")  & ""
	sAction = adoData( "Action")  & ""



	Set xcel = CreateObject("Excel.Application")
	str= Environment("RunFolderPath") & "\" & Environment("ResFldrName") & "\" & DictionaryTest_G("ReportFile")

'	str="C:\Ankit 1\suspenseAccountPaymentsReport\Suspense_account_payments_report_20160330_20160330104839_01.xls"

	Set workbook1 = xcel.Workbooks.Open (str)



'	Set mysheet = workbook1.Worksheets("VF")

	Set mysheet = workbook1.Worksheets("Suspense_account_payments")

	sValue = mysheet.Range(sCell).Value

	If sAction="Compare" Then
		If DictionaryTest_G.Exists(sResult) Then
			sResult = Replace (sResult,sResult,DictionaryTest_G(sResult))
		End If
		If Cstr(sValue)=Cstr(sResult) Then
			AddVerificationInfoToResult  "Pass" , "Value is " & sResult & " as expected"
		else
			AddVerificationInfoToResult  "Fail" , "Actual value is " & sValue & " but expected is " & sResult
			iModuleFail = 1
		End If

	elseif sAction="Capture" Then
		DictionaryTest_G(sResult)=sValue
		AddVerificationInfoToResult  "Pass" , "DictionaryTest_G key " & sResult &" Value is " & sValue
	End If

	workbook1.Close False
	xcel.Quit
	Set xcel = Nothing

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function


'#################################################################################################
' 	Function Name : CompareXcelReportFormat_fn
' 	Description : T
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function CompareXcelReportFormat_fn()

    strSQL = "SELECT * FROM [CompareXcelReportFormat$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\PSCP.xls",strSQL)
	
	'sOR
	sRows1 = adoData( "Rows")  & ""
	sCols1 = adoData( "Cols")  & ""
	sSampleReportPath = adoData( "SampleReportPath")  & ""
	sWorksheet = adoData( "WorkSheet")  & ""

	Set xcelSample = CreateObject("Excel.Application")
	Set xcel = CreateObject("Excel.Application")
	i1=instrrev(Environment("RunFolderPath"),"\")
	fldrPath=Mid(Environment("RunFolderPath"),1,i1-1)
	samplePath = fldrPath &"\SampleReports\" & sSampleReportPath
'
'	Set workbookSample = xcel.Workbooks.Open (samplePath)
'
'	xcel.Application.Visible = False
'
''	Set mysheet = workbook1.Worksheets("VF")
'
'	Set mysheetSample = workbookSample.Worksheets("Suspense_account_payments")
'	str= Environment("RunFolderPath") & "\" & Environment("ResFldrName") & "\" & DictionaryTest_G("ReportFile")



	Set workbook1 = xcel.Workbooks.Open (samplePath)

	xcel.Application.Visible = False

'	Set mysheet = workbook1.Worksheets("VF")


	Set mysheet = workbook1.Worksheets(sWorksheet)
	rows = mysheet.UsedRange.Rows.Count
	cols = mysheet.UsedRange.columns.Count

	str= Environment("RunFolderPath") & "\" & Environment("ResFldrName") & "\" & DictionaryTest_G("ReportFile")
'	str="C:\Ankit 1\suspenseAccountPaymentsReport\Suspense_account_payments_report_20160330_20160330104839_01.xls"
	Set workbook2 = xcelSample.Workbooks.Open (str)

	xcelSample.Application.Visible = False

'	Set mysheet = workbook1.Worksheets("VF")

	Set mysheet2 = workbook2.Worksheets(sWorksheet)

	If sRows1<>"" Then
		
		cellValue1=""
		cellValue2=""
		sRows = Cint(Split(sRows1,"^")(0))
		i = Cint(Split(sRows1,"^")(1))
		
		Do
			cellValue1 = mysheet2.Cells(sRows,i).Value
			cellValue2 = mysheet.Cells(sRows,i).Value
	
			If cellValue1 = cellValue2 Then
				AddVerificationInfoToResult  "Cell " & sRows & "," & i , cellValue1
			else
				AddVerificationInfoToResult  "Fail" , "Actual value is " & cellValue1 & " but expected is " & cellValue2
				iModuleFail = 1
			End If
			i = i+1
		Loop Until (cellValue1 = "" and cellValue1 = "" and i > cols)

	End If

	If sCols1 <> "" Then

		cellValue1=""
		cellValue2=""
		sCols = Cint(Split(sCols1,"^")(0))
		i = Cint(Split(sCols1,"^")(1))
	
		Do
			cellValue1 = mysheet2.Cells(i,sCols).Value
			cellValue2 = mysheet.Cells(i,sCols).Value
	
			If cellValue1 = cellValue2 Then
				AddVerificationInfoToResult  "Cell " & i & "," & sCols , cellValue1
			else
				AddVerificationInfoToResult  "Fail" , "Actual value is " & cellValue1 & " but expected is " & cellValue2
				iModuleFail = 1
			End If
			i = i+1
		Loop Until (cellValue1 = "" and cellValue1 = "" and i > rows)
	
	End If

	workbook2.Close False
	workbook1.Close False
	xcel.Quit
	xcelSample.Quit

	Set xcel = Nothing
	Set xcelSample = Nothing

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function

'#################################################################################################
'               Function Name : ValidateDetRevRaymentRep_fn
'               Description : 
'   Created By :  Ankit
'               Creation Date :        
'##################################################################################################
Public Function ValidateDetRevRaymentRep_fn()

    strSQL = "SELECT * FROM [ValidateDetRevRaymentRep$] WHERE  [RowNo]=" & iRowNo
                Set adoData = ExecExcQry(sProductDir & "Data\PSCP.xls",strSQL)

    sAccountNo = adoData( "AccountNo")  & ""

                If DictionaryTest_G.Exists(sAccountNo) Then
                                sAccountNo = Replace (sAccountNo,sAccountNo,DictionaryTest_G(sAccountNo))
                End If

                sPaymentType = adoData( "PaymentType")  & ""
                sCardType = adoData( "CardType")  & ""

                Set xcel = CreateObject("Excel.Application")

    str= Environment("RunFolderPath") & "\" & Environment("ResFldrName") & "\" & DictionaryTest_G("ReportFile")
'               str="C:\Ankit 1\suspenseAccountPaymentsReport\Suspense_account_payments_report_20160330_20160330104839_01.xls"
                Set workbook2 = xcel.Workbooks.Open (str)

                xcel.Application.Visible = False

'               Set mysheet = workbook1.Worksheets("VF")

                Set mysheet2 = workbook2.Worksheets("VF")

    rows = mysheet2.UsedRange.Rows.Count

                For i = 2 to rows

                                If Cstr(mysheet2.Cells(i,4).Value) =  Cstr(sAccountNo) Then
                                                sPaymentDate = Cstr(DictionaryTest_G("PAYMENTDATE0"))
                                                If instr(sPaymentDate,Cstr(mysheet2.Cells(i,1).Value)) > 0 Then
                                                                                
                                                else
                                                End If

                                                sReversalDate = Cstr(DictionaryTest_G("REVERSALDATE0"))
                                                If instr(sReversalDate,Cstr(mysheet2.Cells(i,2).Value)) > 0 Then
                                                                
                                                else
                                                End If

                                                sAmount = Cstr(DictionaryTest_G("AMOUNT0"))
                                                If instr(sAmount,Cstr(mysheet2.Cells(i,3).Value)) > 0 Then
                                                                
                                                else
                                                End If

                                                If sPaymentType = mysheet2.Cells(i,5).Value Then
                                                                
                                                else
                                                End If

                                                sCusName = Cstr(DictionaryTest_G("CUSNAME0"))
                                                If sCusName = mysheet2.Cells(i,6).Value Then
                                                                
                                                else
                                                End If

                                                sPayTransId = Cstr(DictionaryTest_G("PAYTRANSID0"))
                                                If sCusName = mysheet2.Cells(i,7).Value Then
                                                                
                                                else
                                                End If

                                                If sCardType = mysheet2.Cells(i,10).Value Then
                                                                
                                                else
                                                End If

                                                If mysheet2.Cells(i,12).Value <> "" Then
                                                                
                                                else
                                                End If

                                                If mysheet2.Cells(i,13).Value <> "" Then
                                                                
                                                else
                                                End If

                                                If mysheet2.Cells(i,14).Value <> "" Then
                                                                
                                                else
                                                End If

                                                If mysheet2.Cells(i,15).Value <> "" Then
                                                                
                                                else
                                                End If

                                End If

                Next

                If Err.Number <> 0 Then
                                iModuleFail = 1
                                AddVerificationInfoToResult  "Error" , Err.Description
                End If

End Function




'#################################################################################################
' 	Function Name : DailyCCRefundRep_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function DailyCCRefundRep_fn()

    strSQL = "SELECT * FROM [DailyCCRefundRep$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\PSCP.xls",strSQL)

    sAccountNo = adoData( "AccountNo")  & ""
	If DictionaryTest_G.Exists(sAccountNo) Then
		sAccountNo = Replace (sAccountNo,sAccountNo,DictionaryTest_G(sAccountNo))
	End If

	sCCToken = adoData( "CCToken")  & ""
	sCCExpDate = adoData( "CCExpDate")  & ""
	sCardType = adoData( "CardType")  & ""
	sTotRefAmt = adoData( "TotRefAmt")  & ""




	Set xcel = CreateObject("Excel.Application")

    str= Environment("RunFolderPath") & "\" & Environment("ResFldrName") & "\" & DictionaryTest_G("ReportFile")
'	str="C:\Ankit 1\suspenseAccountPaymentsReport\Suspense_account_payments_report_20160330_20160330104839_01.xls"
	Set workbook2 = xcel.Workbooks.Open (str)

	xcel.Application.Visible = False

'	Set mysheet = workbook1.Worksheets("VF")

	Set mysheet2 = workbook2.Worksheets("VF")

    rows = mysheet2.UsedRange.Rows.Count

	flag = "N"
		
	For i = 2 to rows

		If Cstr(mysheet2.Cells(i,5).Value) =  Cstr(sAccountNo) Then
			sPaymentDate = Cstr(DictionaryTest_G("PAYMENTDATE0"))
			If instr(sPaymentDate,Cstr(mysheet2.Cells(i,1).Value)) > 0 Then
					
			else
			End If

			sRefundDate = Cstr(DictionaryTest_G("REFUNDDATE0"))
			If instr(sRefundDate,Cstr(mysheet2.Cells(i,2).Value)) > 0 Then
				
			else
			End If

			sRefund = Cstr(DictionaryTest_G("Refund"))
			If instr(sRefund,Cstr(mysheet2.Cells(i,3).Value)) > 0 Then
				
			else
			End If

			sAmount = Cstr(DictionaryTest_G("AMOUNT0"))
			If instr(sAmount,Cstr(mysheet2.Cells(i,4).Value)) > 0 Then
							
			else
			End If

			sCusName = Cstr(DictionaryTest_G("CUSNAME0"))
			If sCusName = mysheet2.Cells(i,6).Value Then
				
			else
			End If

			
			If sCCToken = mysheet2.Cells(i,7).Value Then
				
			else
			End If

			If sCCExpDate = mysheet2.Cells(i,8).Value Then
				
			else
			End If

			If sCardType = mysheet2.Cells(i,9).Value Then
				
			else
			End If
			Exit For
			flag = "Y"
			
		End If

	Next

	If flag = "N" Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "Account no not found"
	End If

	For j = i to rows
		If instr(Cstr(mysheet2.Cells(j,3).Value),Cstr("Total refund amount:")) > 0 Then
			If sTotRefAmt = "Capture" Then
				DictionaryTest_G ("TotRefAmt") = Cstr(mysheet2.Cells(j+1,3).Value)
			elseIf sTotRefAmt = "Compare" Then
				If Cstr(mysheet2.Cells(j+1,3).Value) >  DictionaryTest_G ("TotRefAmt") Then
					AddVerificationInfoToResult  "Pass" , "Amount refunded"
				else
					AddVerificationInfoToResult  "Fail" , "Amount not refunded"
					iModuleFail = 1
				End If
			End If
		else
			iModuleFail = 1
			AddVerificationInfoToResult  "Error" , "Total refund amount is not present in report"
		End if
	Next

	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function


'#################################################################################################
' 	Function Name : RevPaymentRep_fn
' 	Description : 
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function RevPaymentRep_fn()

    strSQL = "SELECT * FROM [RevPaymentRep$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\PSCP.xls",strSQL)


	sPaymentMethod = adoData( "PaymentMethod")  & ""
	sTotRefAmt = adoData( "TotRefAmt")  & ""
	sKey = adoData( "Key")  & ""


	Set xcel = CreateObject("Excel.Application")

    str= Environment("RunFolderPath") & "\" & Environment("ResFldrName") & "\" & DictionaryTest_G("ReportFile")
'	str="C:\VATS_Automation_BRM-SVN\Automation\Driver\Results\BRM\10122016\Pay11_j78ex4nnpg\Reversed_payments_report.xls"
	Set workbook2 = xcel.Workbooks.Open (str)

	xcel.Application.Visible = False

'	Set mysheet = workbook1.Worksheets("VF")

	Set mysheet2 = workbook2.Worksheets("VF")

    rows = mysheet2.UsedRange.Rows.Count

	flag = "N"
		
	For i = 2 to rows

		If Cstr(mysheet2.Cells(i,1).Value) =  Cstr(sPaymentMethod) Then
			If sTotRefAmt = "Capture" Then
				DictionaryTest_G (sKey) = Cstr(mysheet2.Cells(i,2).Value)
				AddVerificationInfoToResult  "Info" , "for Payment method" & sPaymentMethod  & "" & "the amount captured is" & DictionaryTest_G ("TotRefAmt")
			elseIf sTotRefAmt = "Compare" Then
				If Cstr(mysheet2.Cells(i,2).Value) >= Cstr(DictionaryTest_G (sKey) + DictionaryTest_G ("AMOUNT0"))Then
					AddVerificationInfoToResult  "Pass" , "Amount reflected correctly"
				else
					AddVerificationInfoToResult  "Fail" , "Amount not reflected correctly"
					iModuleFail = 1
				End If
			End If
		End If

	Next


	If Err.Number <> 0 Then
		iModuleFail = 1
		AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function



'#################################################################################################
' 	Function Name : Printshop_fn
' 	Description : 
'   Created By :  Anujita
'	Creation Date :        
'##################################################################################################
Public Function Printshop_fn()

   'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [Printshop$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\PSCP.xls",strSQL)

	sFileType = adoData("ReportName") & ""
	iIndex = adoData("Index") & ""

	If iIndex="" Then
		iIndex=0
	End If

	call SetObjRepository ("Putty",sProductDir & "Resources\")
	Window("PuTTY").Activate
	Window("PuTTY").PuttyType  "$","cd /opt/odi1/Middleware/Oracle_ODI1/DocStore/FileMerged"

	Window("PuTTY").PuttyType  "$","time=`date +%Y%m%d`"
	'Window("PuTTY").PuttyType  "$","time=20161014"
	'Window("PuTTY").PuttyType  "$","ls " & "VF_" & chr(34) & "$time" & chr(34) & "_ODI_BL_UK_1"  & ".zip"
	Window("PuTTY").PuttyType  "$","ls " & "VF_" & chr(34) & "$time" & chr(34) & "_ODI_" & sFileType & "_*" & ".zip"

	iIndex=0

	sValidatePutty = ValidatePutty ("ls VF_ ","VF_")
	CaptureSnapshot()
	If sValidatePutty = True Then
		AddVerificationInfoToResult  "Info" , "Printshop File generated at FMW"
	Else
		AddVerificationInfoToResult  "Error" , "Printshop file could not generated"
		iModuleFail = 1
		Exit Function
	End If

	'Window("PuTTY").PuttyType  "$","zipfile=" & chr(96) & "ls " & "VF_" & chr(34) & "$time" & chr(34) & "_ODI_" & "sFileType" & ".zip" & chr(96)
	'Window("PuTTY").PuttyType  "$","chmod 777 " & "VF_" & chr(34) & "$time" & chr(34) & "_ODI_" & "sFileType" & ".zip"
	
	sCapturePuttyData = CapturePuttyData ("LINEAFTER:ls VF_","ReportFile",iIndex)
	If Not(sCapturePuttyData = True) Then
		 iModuleFail = 1
		AddVerificationInfoToResult  "Error" , "ReportFile is not present"
		Exit Function
	End If

	sReportFile = DictionaryTest_G("ReportFile")
	
	DictionaryTest_G("ReportFile") = "VF_" & Split(Split(sReportFile,"VF_")(1),".zip")(0) & ".zip"
	Window("PuTTY").PuttyType  "$","chmod 777 " & DictionaryTest_G("ReportFile")
	'DictionaryTest_G ("ReportFile") = Replace(DictionaryTest_G ("ReportFile")," ","")
	'DictionaryTest_G ("ReportFile")="VF_20161014_ODI_BL_UK_1.zip"

	CaptureSnapshot()

	If Err.Number <> 0 Then
            iModuleFail = 1
			AddVerificationInfoToResult  "Error" , Err.Description
	End If

End Function
