Option Explicit
Dim strExcelPath, oXL, oXLBook, oXLSheet, bReadOnly, objShell, intExcelRow, intExcelColumn, doubleExcelInputValue, strSheetName, intExcelNoColumns, intExcelNoRows, i, j

' Read values from evoware
strExcelPath = cstr(Evoware.GetStringVariable("VBScript_SetExcelSheetMultipleCellValueNumber_FileNamePath"))
intExcelRow = cint(Evoware.GetDoubleVariable("VBScript_SetExcelSheetMultipleCellValueNumber_Row"))
intExcelColumn = cint(Evoware.GetDoubleVariable("VBScript_SetExcelSheetMultipleCellValueNumber_Column"))
intExcelNoColumns = cint(Evoware.GetDoubleVariable("VBScript_SetExcelSheetMultipleCellValueNumber_NoColumns"))
intExcelNoRows = cint(Evoware.GetDoubleVariable("VBScript_SetExcelSheetMultipleCellValueNumber_NoRows"))
doubleExcelInputValue = Evoware.GetDoubleVariable("VBScript_SetExcelSheetMultipleCellValueNumber_Number")
strSheetName = cstr(Evoware.GetStringVariable("VBScript_SetExcelSheetMultipleCellValueNumber_SheetName"))

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
	objShell.Popup "The excel file is currently in read-only mode. I will try again in 2 second" , 2
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
Set oXLSheet = oXL.ActiveWorkbook.Worksheets(strSheetName)
'Activate worksheet (same as clicking tab)
oXLSheet.Activate
'Write values to worksheet
For i = 0 to intExcelNoColumns-1
	For j = 0 to intExcelNoRows-1
    oXLSheet.Cells(intExcelRow+j, intExcelColumn+i).Value = doubleExcelInputValue
	Next
Next
'Save and quit.
oXL.ActiveWorkbook.Save
oXLBook.Close False
Set oXLBook = Nothing
Set oXLSheet = Nothing

oXL.Quit
'Housekeeping
Set oXL = Nothing		
Set bReadOnly = Nothing
Set strExcelPath = Nothing
Set intExcelRow = Nothing
Set intExcelColumn = Nothing
Set doubleExcelInputValue = Nothing
Set strSheetName = Nothing

Sub Delay(seconds) 'Submethod for delay without using Wscript
	Dim wshShell, strCmd
	Set wshShell = CreateObject( "WScript.Shell" )
	strCmd = wshShell.ExpandEnvironmentStrings( "%COMSPEC% /C (TIMEOUT.EXE /T " & seconds & " /NOBREAK)" )
	wshShell.Run strCmd, 0, 1
	Set wshShell = Nothing
End Sub