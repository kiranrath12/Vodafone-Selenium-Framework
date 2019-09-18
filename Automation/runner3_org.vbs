On error resume next
Set objShell = CreateObject("WScript.Shell")
REM objShell.Exec "taskkill /F /IM EXCEL.EXE"
REM objShell.Exec "taskkill /F /IM QTPro.exe *32"
REM objShell.Exec "taskkill /F /IM QTAUTOMATIONAGENT.EXE"
REM Set objProcs = GetObject("winmgmts:").ExecQuery("select * from Win32_Process where Name = 'QTPro.exe' OR Name =  'QTAutomationAgent.exe' or Name = 'EXCEL.exe'")
    REM For Each Process In objProcs
        REM Process.Terminate
    REM Next
REM wait 10
Set objEnv = objShell.Environment("User")
objEnv("Module") = WScript.Arguments.Item(0)
objEnv("Environment") = WScript.Arguments.Item(1)
objEnv("TCResult") = ""
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
If err.number <> 0 or TCResults = "" Then
WScript.StdOut.WriteLine err.description
	WScript.Quit(2)
else
	WScript.Quit(0)
End If
