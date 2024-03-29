'#################################################################################################
'               Function Name : QCUpload_fn
'               Description : This function Zips the test case results and uploads them to the respective path in QC
'   Created By :  Pushkar Vasekar
'               Creation Date :        
'##################################################################################################'''

Public Function QCUpload_fn
                Dim adoData        
    strSQL = "SELECT * FROM [QCCredentials$] WHERE  [RowNo]=" & iRowNo
                Set adoData = ExecExcQry(sProductDir & "Data\QC.xls",strSQL)

                QCUserName = adoData("QCUname") & ""
                QCPassword = adoData("QCPassword") & ""
                sDomain = adoData("Domain") & ""
                sProjectName = adoData("ProjectName") & ""
                sServerName = adoData("ServerName") & ""
                TestCaseName = sTestCaseName 
                QCFolderPath =sQCFolderPath
                TestSetName =sQCTestSetName
                TCStatus = "Passed"


                ResPath =Environment("RunFolderPath") & "\" & Environment("ResFldrName") 
                ZipPath = Environment("RunFolderPath") & "\" & Environment("ResFldrName")  & ".zip"
                Call ZipFolder (ResPath,ZipPath)
				If Err.Number <> 0 or iModuleFail = 1 Then
					
					AddVerificationInfoToResult  "Error" , "Error in Zip folder function"
					Exit Function
				End If
                Call QC_Status_Upload(QCUserName, QCPassword,QCFolderPath,TestSetName,TestCaseName,TCStatus,ZipPath,sDomain,sProjectName,sServerName)
				If Err.Number <> 0 or iModuleFail = 1 Then
					
					AddVerificationInfoToResult  "Error" , "Error in QC_Status_Upload function"
					Exit Function
				End If
End Function

Dim iRow, QCUserName, QCPassword,QCFolderPath,TestSetName,TestCaseName,TCStatus,ResPath,myZipFile


''''''''''''''''''''+++++++++++++++++++++++++++++++++++++++++++++++END FUNCTION+++++++++++++++++++++++++++++++++++++++++++++++

                                

''''+++++++++++++++++++Function to ZIP the result file++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function ZipFolder(  myFolder, myZipFile )

    Dim objApp, objFolder, objFSO, objItem, objTxt
    Dim strSkipped
    AddVerificationInfoToResult "Info" ,"In Zip function"
    Const ForWriting = 2
    intSkipped = 0
    ' Make sure the path ends with a backslash
    If Right( myFolder, 1 ) <> "\" Then
        myFolder = myFolder & "\"
    End If
    ' Use custom error handling
    On Error Resume Next
    ' Create an empty ZIP file
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

If  objFSO.FileExists(myZipFile)Then
        objFSO.DeleteFile myZipFile, True
End If

    Set objTxt = objFSO.OpenTextFile( myZipFile, ForWriting, True )
    objTxt.Write "PK" & Chr(5) & Chr(6) & String( 18, Chr(0) )
    objTxt.Close
    Set objTxt = Nothing
    ' Abort on errors
    If Err Then
		AddVerificationInfoToResult "Error" , "Error in Zipping file"
		iModuleFail = 1
        Exit Function
	else
		AddVerificationInfoToResult "Info" ,"Zip folder created"
    End If
    
    ' Create a Shell object
    Set objApp = CreateObject( "Shell.Application" )
    ' Copy the files to the compressed folder
    For Each objItem in objApp.NameSpace( myFolder ).Items
'        If objItem.IsFolder Then
'            ' Check if the subfolder is empty, and if
'            ' so, skip it to prevent an error message
'            Set objFolder = objFSO.GetFolder( objItem.Path )
'            If objFolder.Files.Count + objFolder.SubFolders.Count = 0 Then
'                intSkipped = intSkipped + 1
'            Else
'                objApp.NameSpace( myZipFile ).CopyHere objItem
'            End If
'        Else
		
        objApp.NameSpace( myZipFile ).CopyHere objItem
		wait 1
'        End If
    Next
    Set objFolder = Nothing
    Set objFSO    = Nothing
    ' Abort on errors
    If  objApp.NameSpace( myFolder ).Items.Count=0 Then
		AddVerificationInfoToResult "Error" , "Zip Folder is empty."
		iModuleFail = 1
        Exit Function
	else
		AddVerificationInfoToResult "Info" , "Files copied to zip folder"
    End If

	If objApp.NameSpace( myZipFile ).Items.Count<> objApp.NameSpace( myFolder ).Items.Count Then
		AddVerificationInfoToResult "Error" , "All files not copied into zip folder."
		iModuleFail = 1
        Exit Function
	End If
    ' Keep script waiting until compression is done
'	On Error Resume Next
'    intSrcItems = objApp.NameSpace( myFolder  ).Items.Count
'    Do Until objApp.NameSpace( myZipFile ).Items.Count + intSkipped = intSrcItems
'        WScript.Sleep 10
'    Loop
    Set objApp = Nothing
'	Err.clear
    ' Abort on errors
    If Err Then
		AddVerificationInfoToResult "Error" , "Error"
		iModuleFail = 1
        Exit Function
	else
		AddVerificationInfoToResult "Info" ,"Result for " &TestCaseName& " under " &TestSetName& " is zipped successfully."
    End If

    ' Return message if empty subfolders were skipped
'    If intSkipped = 0 Then
'        strSkipped = ""
'    Else
'        strSkipped = "skipped empty subfolders"
'    End If
    
'    ZipFolder = Array( 0, intSkipped, strSkipped )
                
    
   
                
End Function


''''''''''''''''''++++++++++++++++++++++++++++Function for QC attachment upload and status update+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function QC_Status_Upload(QCUserName, QCPassword,QCFolderPath,TestSetName,TestCaseName,TCStatus, ZipPath,sDomain,sProjectName,sServerName)

                qcUser = QCUserName
   qcPassword = QCPassword
   qcDomain = sDomain
   qcProject =sProjectName
   udv_qcServer =sServerName


   Set udv_tdc = CreateObject("TDApiOle80.TDConnection")
   'Initialise the Quality center connection
   udv_tdc.InitConnectionEx udv_qcServer

   'Loging the username and password
   udv_tdc.Login qcUser, qcPassword

   'connecting to the domain and project
   udv_tdc.Connect qcDomain, qcProject
                AddVerificationInfoToResult "Info" ,"QCconnected"

    folder_path= QCFolderPath ' the path of the folder which contains your test set

'  For each fil in resFolder.Files
   Set TSetFact = udv_tdc.TestSetFactory

  Set tsTreeMgr = udv_tdc.TestSetTreeManager

   Set TestSetFolder1=tsTreeMgr.NodeByPath(folder_path)   

   set tsff = TestSetFolder1.TestSetFactory.Filter

   tsff.Filter("CY_CYCLE")= TestSetName
AddVerificationInfoToResult "Info" ,"In QC  function"
   set tsl = tsff.NewList()

   If tsl.count = 0 Then
       AddVerificationInfoToResult "Info" , "Test set with the name " & TestSetName & " is not found "
                                iModuleFail = 1
     ' Print "Test set with the name " & TestSetName & " is not found "
   End If

   For i=1 to tsl.count
      If  tsl.Item(i).Name = sQCTestSetName  Then

      Set ts = tsl.Item(i)
      Set test_case_filter = ts.TSTestFactory.Filter
      filter_string= Chr(34) & "*" & Trim(TestCaseName) & "*" &Chr(34)
                                'filter_string = test_name
     test_case_filter.Filter("TS_NAME")=TRIM(filter_string)
      Set tcase=test_case_filter.NewList()

      If tcase.count = 0 Then
                                  AddVerificationInfoToResult "Info" , "Test Case  with the name " & TestCaseName & " is not found "
                                  iModuleFail = 1
'         Print "Test case with the name " & TestCaseName & " is not found "
      End If

      For j=1 to tcase.count

                                tcase(j).Field("TC_STATUS")= TCStatus
         tcase(j).Post

                                  Set strattachPath =tcase(j).Attachments
                                  Set strFile = strattachPath.AddItem(Null)
                                  Set objFile = CreateObject("Scripting.FileSystemObject")

                                                If objFile.FileExists(ZipPath) Then
                                                strFile.FileName = ZipPath
                                                strFile.Type = 1
                                                strFile.Description = " HTML Report"
                                                strFile.Post
                                                End If
                                'Next
       
      Next   ' End of test CASE loop
End If
   Next   'End of test SET loop

                                AddVerificationInfoToResult "Info" ,"Result for " &TestCaseName& " uploaded to QC  successfully and status is Passed."
    

End Function

''''''''''''''''''''+++++++++++++++++++++++++++++++++++++++++++++++END FUNCTION+++++++++++++++++++++++++++++++++++++++++++++++
   Set udv_tdc = nothing
   Set TSetFact = nothing
   Set TestSetFolder1 = nothing


''''''''''''''''''''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++