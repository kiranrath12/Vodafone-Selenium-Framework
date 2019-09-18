Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment("User")
objEnv("Module") = WScript.Arguments.Item(0)
objEnv("Environment") = WScript.Arguments.Item(1)
'msgbox objEnv("Module")
'msgbox objEnv("Environment")
Call Execute_Click()
TCResults=objEnv("TCResult")
arrTCResults=Split(TCResults,";")
WScript.StdOut.WriteLine "Start"
For each arrTCResult in arrTCResults
	'TCName = Split(arrTCResult)(0)
	'Result = Split(arrTCResult)(1)
	WScript.StdOut.WriteLine arrTCResult
Next
WScript.StdOut.WriteLine "End"
If err.number = 0 or instr(TCResults,"SBL")>0 Then	
	WScript.Quit(0)
else
	WScript.StdOut.WriteLine err.description
	WScript.Quit(2)	
End If



Dim arrVariable
Dim qtApp
Dim qtTest
Dim qtResultsOpt
Dim QueryCount
Dim sDriverPath
Dim sLogFilePath
Dim runFolder
Dim sJourneyType
Dim sXMLResultFile
Dim  Env
'Env ="E4"


Sub Execute_Click()
'    Call GatherQueries
    Dim sExecutionType
    Dim sQtpTestStatus
'    sExecutionType = MsgBox("Click Yes to Start Execution OR No to Launch Driver (for Debugging)", vbYesNoCancel)
'    If sExecutionType = 2 Then
'        Exit Sub
'    End If
    sDriverPath = "C:\VATS Automation Sanity\Automation\Driver - Jenkins"
    Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions") ' Create the Run Results Options object
    
    

    sLogFilePath = sDriverPath & "\DriverLog_" & Replace(Replace(Now, ":", "-"), "/", "-") & ".txt"
    Call CreateTextFile(sLogFilePath)
    Set qtApp = CreateObject("QuickTest.Application")
    qtApp.Visible = False
    If IsQTPApplicationStarted(qtApp) = "No" Then
        qtApp.launch
        
    End If
    If qtApp.GetStatus = "Ready" Then
        qtApp.Open sDriverPath, False, False
        WScript.Sleep 10
    End If
    qtApp.WindowState = "Maximized"
    Set qtTest = qtApp.Test
	Dim arrQuery(1,5)
	arrQuery(0, 0) = "NewCo"
	arrQuery(0, 1) = "SQL"
	arrQuery(0, 2) = "C:\VATS Automation Sanity\Automation\NewCo\Data\Testcases_Bars.xls"
	arrQuery(0, 3) = "SELECT * FROM [Testcases$] WHERE  [Type] = 'Sanity'"
	arrQuery(0, 4) = "Sanity"
    
'    Call GatherQueries
    WScript.Sleep 2
    runFolder = CreateResultsFolder(sDriverPath & "\Results\" & arrQuery(0, 0))
    qtTest.Environment.Value("RunFolderPath") = runFolder
    Dim objShell, objEnv
    Set objShell = CreateObject("WScript.Shell")
    Set objEnv = objShell.Environment("User")
 
    objEnv("outputRowNum") = 0
    'qtTest.Environment.Value("outputRowNum") = 0
    Dim OutputFile
    OutputFile = CreateOutputXcelFile(sDriverPath & "\Results\" & arrQuery(0, 0), runFolder)
    qtTest.Environment.Value("OutputFile") = OutputFile
'    For QueryCount = 0 To GatherQueries - 1
        Call WriteLog(sLogFilePath, Now & "--" & "Query" & "--" & arrQuery(0, 3))
        
        
'        If qtTest.Environment.Value("Product") = arrQuery(0, 0) Then
            'Do not reload the libraries
'        Else
            'Load libraries for this product
            qtTest.Environment.Value("Product") = arrQuery(0, 0)
            qtTest.Environment.Value("Module") = arrQuery(0, 4)
           ' Call LoadLib(qtApp, qtTest.Environment.Value("Product"))
'        End If
        qtTest.Environment.Value("VariableSet") = arrQuery(0, 1)
        qtTest.Environment.Value("TestCaseFile") = arrQuery(0, 2)
        qtTest.Environment.Value("Query") = arrQuery(0, 3)
        sJourneyType = ""
			sXMLResultFile = ""
			runFolder = ""

			arrVariable = "sURL^|sEnv^E7"
'        Call GatherVariables(arrQuery(0, 0), arrQuery(0, 1))
        qtTest.Environment.Value("VariableSet") = arrVariable
        If sJourneyType = "SecondPartOnly" Then
            qtTest.Environment.Value("ResultFileName") = sXMLResultFile
        Else
            qtTest.Environment.Value("ResultFileName") = arrQuery(0, 0) & "_" & arrQuery(0, 4) & "_" & Replace(Replace(Now, ":", "i"), "/", "i")
        End If
        qtApp.Visible = False
        
        'Dim runFolder
       ' runFolder = CreateResultsFolder(sDriverPath & "\Results\" & arrQuery(0, 0))
        'qtTest.Environment.Value("RunFolderPath") = runFolder
        
'        If sExecutionType = 7 Then
'            Exit Sub
'        End If
        
        qtResultsOpt.ResultsLocation = runFolder & "\" & qtTest.Environment.Value("ResultFileName") ' Set the results location
       ' qtResultsOpt.ResultsLocation = sDriverPath & "\Results\" & arrQuery(0, 0) & "\" & qtTest.Environment.Value("ResultFileName") ' Set the results location
        On Error Resume Next
		WScript.Sleep 5
'        Application.WScript.Sleep (Now + TimeValue("0:00:05"))
        qtTest.Run qtResultsOpt, False
		WScript.Sleep 50
'        Application.WScript.Sleep (Now + TimeValue("0:00:50"))
        Do While qtApp.Test.IsRunning
        '    sQtpTestStatus = QtpTestStatus(qtApp)
         '   If sQtpTestStatus = "Stopped" Or sQtpTestStatus = "Hang" Then
          '      Call WriteLog(sLogFilePath, Now & "--" & sQtpTestStatus)
           '     Exit Do
            'End If
		Loop
'    Next
End Sub

Public Function IsQTPApplicationStarted(qtApp)
    On Error Resume Next
    qtApp.Test.IsRunning

    If Err.Description = "The operation failed because the application was not launched" Then
        IsQTPApplicationStarted = "No"
    Else
        IsQTPApplicationStarted = "Yes"
    End If
End Function
Public Function CreateTextFile(sLogFilePath)
    Dim fso, OutputFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set OutputFile = fso.CreateTextFile(sLogFilePath, True)
    OutputFile.Close
End Function

Public Function WriteLog(sLogFilePath, sMatter)
    Dim fso, OutputFile
    Const ForAppending = 8
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set OutputFile = fso.OpenTextFile(sLogFilePath, ForAppending, True)
    OutputFile.WriteLine (sMatter)
    OutputFile.Close
End Function

Public Function CreateResultsFolder(sPath)
    Dim fso, f, sDate, path
    Set fso = CreateObject("Scripting.FileSystemObject")
    sDate = Replace(Date, "/", "")
    If (fso.FolderExists(sPath & "\" & sDate)) = False Then
        Set f = fso.CreateFolder(sPath & "\" & sDate)
        path = f.path
    Else
        path = sPath & "\" & sDate
    End If
    fso.CopyFile sPath & "\minus.gif", path & "\", True
    fso.CopyFile sPath & "\plus.gif", path & "\", True
    fso.CopyFile sPath & "\leaf.gif", path & "\", True
    fso.CopyFile sPath & "\Result.dtd", path & "\", True
    fso.CopyFile sPath & "\Result.xsl", path & "\", True
    CreateResultsFolder = path

End Function

Public Function CreateOutputXcelFile1(sPath)
    Dim xcel, outpoutFileName, Worksheet
    Set xcel = CreateObject("Excel.Application")
    outpoutFileName = "OutputSheet_" & Replace(Replace(Now, ":", "-"), "/", "-")
    xcel.Workbooks.Add
    Set Worksheet = xcel.Worksheets("Sheet1")
    Worksheet.Cells(1, 1).Value = "RowNo"
    Worksheet.Cells(1, 2).Value = "Env"
    Worksheet.Cells(1, 3).Value = "Module"
    Worksheet.Cells(1, 4).Value = "TCID"
    Worksheet.Cells(1, 4).Value = "TCName"
    Worksheet.Cells(1, 5).Value = "AccountNo"
    Worksheet.Cells(1, 6).Value = "OrderNo"
    Worksheet.Cells(1, 7).Value = "MSISDN"
    Worksheet.Cells(1, 8).Value = "Status"
    Worksheet.Cells(1, 9).Value = "TCResult"
    Worksheet.Cells(1, 10).Value = "ResFldrName"

    'Worksheet.Cells(1, 10).Value = 1
    
    xcel.ActiveWorkbook.SaveAs (sPath & "\" & outpoutFileName & ".xls")
    xcel.Application.Quit
    Set xcel = Nothing
    CreateOutputXcelFile = outpoutFileName & ".xls"
    

End Function

Public Function CreateOutputXcelFile(copyFrom, copyTo)
    Dim fso, outpoutFileName
    outpoutFileName = "OutputSheet_" & Replace(Replace(Now, ":", "-"), "/", "-")
    Set fso = CreateObject("Scripting.FileSystemObject")

    fso.CopyFile copyFrom & "\OutputSheet.xls", copyTo & "\" & outpoutFileName & ".xls"
    
    CreateOutputXcelFile = outpoutFileName & ".xls"
End Function

