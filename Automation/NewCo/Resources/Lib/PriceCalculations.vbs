'#################################################################################################
' 	Function Name : LoginToPutty_fn
' 	Description : This function log into Siebel Application
'   Created By :  Ankit
'	Creation Date :        
'##################################################################################################
Public Function LoginToPutty_fn()

'Get Data
	Dim adoData	  
    strSQL = "SELECT * FROM [LoginToPutty$] WHERE  [RowNo]=" & iRowNo
	Set adoData = ExecExcQry(sProductDir & "Data\Putty.xls",strSQL)


End Function