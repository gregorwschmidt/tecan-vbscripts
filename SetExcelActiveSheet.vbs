Option Explicit
Dim strExcelPath, strSheetName, oXL, oXLBook, oXLSheet, bReadOnly, objShell

	'Read values from evoware
	strExcelPath = cstr(Evoware.GetStringVariable("VBScript_SetExcelActiveSheet_FileNamePath"))
	strSheetName = cstr(Evoware.GetStringVariable("VBScript_SetExcelActiveSheet_SheetName"))

	
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
	
	While bReadOnly = True
		'Create popup window
		Set objShell = CreateObject("WScript.Shell")
		objShell.Popup "The excel file is currently in read-only mode. I will try again in 5 seconds" , 2
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
	'Save and quit.
	oXL.ActiveWorkbook.Save
	oXLBook.Close False
	Set oXLSheet = Nothing
	Set oXLBook = Nothing
	
	oXL.Quit
	'Housekeeping
	Set oXL = Nothing		
	Set bReadOnly = Nothing
	Set strSheetName = Nothing
	Set strExcelPath = Nothing
	
Sub Delay(seconds) 'Submethod for delay without using Wscript
	Dim wshShell, strCmd
	Set wshShell = CreateObject( "WScript.Shell" )
	strCmd = wshShell.ExpandEnvironmentStrings( "%COMSPEC% /C (TIMEOUT.EXE /T " & seconds & " /NOBREAK)" )
	wshShell.Run strCmd, 0, 1
	Set wshShell = Nothing
End Sub