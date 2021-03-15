Option Explicit
Dim strExcelPath, oXL, oXLBook, oXLSheet, bReadOnly, objShell, strExcelInputValue, strExcelVariableName

' Read values from evoware
strExcelPath = cstr(Evoware.GetStringVariable("VBScript_SetExcelVariableValueString_FileNamePath"))
strExcelVariableName = cstr(Evoware.GetStringVariable("VBScript_SetExcelVariableValueString_VariableName"))
strExcelInputValue = cstr(Evoware.GetStringVariable("VBScript_SetExcelVariableValueString_String"))


'Open Excel and disable alerts
Set oXL =  CreateObject("Excel.Application")
'Wait one second for excel to properly open
Delay 1
oXL.DisplayAlerts = False
'Open workbook and activate it
Set oXLBook = oXL.Workbooks.Open(strExcelPath)
oXLBook.Activate
'Check if read only status applies
bReadOnly = oXL.ActiveWorkBook.ReadOnly

'While sheet in read-only mode, display a short popup window, close the file and reopen it and check status again
While bReadOnly = True
	'Create popup window
	Set objShell = CreateObject("WScript.Shell")
	objShell.Popup "The excel file is currently in read-only mode. I will try again in 2 seconds" , 2
	Set objShell = Nothing
	'Close workbook and end excel application
	oXLBook.Close False
	Set oXLBook = Nothing
	
	oXL.Quit
	Set oXL = Nothing
	
	Randomize
	If Rnd > 0.5 Then 'Add stochastic component such that robots will not block each other when trying to access the files
		Delay 1
	Else
		Delay 3
	End If 
	
	'Open Excel again and disable alerts
	Set oXL =  CreateObject("Excel.Application")
	oXL.DisplayAlerts = False
	'Open workbook again and activate it
	Set oXLBook = oXL.Workbooks.Open(strExcelPath)
	oXLBook.Activate
	'Check if read only status still applies
	bReadOnly = oXL.ActiveWorkBook.ReadOnly	
Wend
	
'Choose desired worksheet
Set oXLSheet = oXL.ActiveWorkbook.Worksheets(1)
'Write value to named variable in worksheet
oXLSheet.Range(strExcelVariableName).Value = strExcelInputValue
'Save and quit.
oXL.ActiveWorkbook.Save
oXLBook.Close False
Set oXLSheet = Nothing
Set oXLBook = Nothing

oXL.Quit
'Housekeeping
Set oXL = Nothing		
Set bReadOnly = Nothing
Set strExcelPath = Nothing
Set strExcelVariableName = Nothing
Set strExcelInputValue = Nothing


Sub Delay(seconds) 'Submethod for delay without using Wscript
	Dim wshShell, strCmd
	Set wshShell = CreateObject( "WScript.Shell" )
	strCmd = wshShell.ExpandEnvironmentStrings( "%COMSPEC% /C (TIMEOUT.EXE /T " & seconds & " /NOBREAK)" )
	wshShell.Run strCmd, 0, 1
	Set wshShell = Nothing
End Sub