Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment("User")
objEnv("Module") = WScript.Arguments.Item(0)
objEnv("Environment") = WScript.Arguments.Item(1)
'msgbox objEnv("Module")
'msgbox objEnv("Environment")
Set xl = CreateObject("Excel.Application")
Set wb = xl.Workbooks.Open("C:\VATS Automation Sanity\Automation\NewCo\StartDriver - Jenkins.xlsm")
xl.Run "Sheet5.Execute_Click"
wb.Close
xl.Quit
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
