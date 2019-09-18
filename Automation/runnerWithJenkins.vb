Call StartJenkins
'msgbox WScript.Arguments.Item(0)
Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment("User")
objEnv("Module") = WScript.Arguments.Item(0)
objEnv("Environment") = WScript.Arguments.Item(1)
'msgbox objEnv("Module")
'msgbox objEnv("Environment")
Set xl = CreateObject("Excel.Application")
Set wb = xl.Workbooks.Open("C:\VATS Automation\Automation\NewCo\StartDriver - Jenkins.xlsm")

REM xl.Run "'TestCases'!Execute_Click"
REM Set wb = xl.Workbooks.Open("C:\Sanity Data Bank_1\Automation\NewCo\Book1.xlsm")
xl.Run "Sheet5.Execute_Click"
wb.Close
xl.Quit
TCResults=objEnv("TCResult")
arrTCResults=Split(TCResults,";")
For each arrTCResult in arrTCResults
	'TCName = Split(arrTCResult)(0)
	'Result = Split(arrTCResult)(1)
	WScript.StdOut.WriteLine arrTCResult
Next

Public Function StartJenkins
Set WshShell = CreateObject("wscript.Shell")
WshShell.Run "cmd"
WScript.Sleep 1000
WshShell.AppActivate "C:\Windows\system32\cmd.exe" 
WScript.Sleep 1000
WshShell.sendkeys "Cd C:\Jenkins"
WshShell.sendkeys "{ENTER}"
WshShell.sendkeys "java -jar slave.jar -jnlpUrl http://10.78.192.231:8080/jenkins/computer/Box132/slave-agent.jnlp -secret 98df2be4a8622e1fffa91753ee933a659daa2fd9934373c8eec05eba1c30430a"
WshShell.sendkeys "{ENTER}"
WScript.Sleep 20000
End Function
