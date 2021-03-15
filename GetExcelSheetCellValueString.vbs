Option Explicit
Dim strExcelFileNamePath, oXL, oXLBook, oXLSheet, objShell, intExcelRow, intExcelColumn, strExcelOutputValue, strSheetName
Dim fso, strExcelFilePath, Temp, strFileExtension, strExcelCopyFileNamePath, dateTime

' Read values from evoware
strExcelFileNamePath = cstr(Evoware.GetStringVariable("VBScript_GetExcelSheetCellValueString_FileNamePath"))
intExcelRow = cint(Evoware.GetDoubleVariable("VBScript_GetExcelSheetCellValueString_Row"))
intExcelColumn = cint(Evoware.GetDoubleVariable("VBScript_GetExcelSheetCellValueString_Column"))
strSheetName = cstr(Evoware.GetStringVariable("VBScript_GetExcelSheetCellValueString_SheetName"))

'Create file system object
Set fso = CreateObject("Scripting.FileSystemObject")

'Build file path for temporary excel file 
strExcelFilePath =  fso.GetParentFolderName(strExcelFileNamePath)&"\"
Temp = split(strExcelFileNamePath, ".")
strFileExtension = Temp(1)
Randomize
dateTime = Year(now) & "-" & right("00" & Month(now),2) & "-" & right("00" & Day(now),2) & "_" & right("00" & Hour(now),2) & "-" & right("00" & Minute(now),2) & "-" & right("00" & Second(now),2) & "-" & cstr(int((10000-1+1)*Rnd+1))
strExcelCopyFileNamePath = cstr(strExcelFilePath & "TEMP_" & dateTime & "." & strFileExtension)

'Housekeeping of first part
Set strExcelFilePath = Nothing
Set Temp = Nothing
Set strFileExtension = Nothing
Set dateTime = Nothing

If fso.FileExists(strExcelFileNamePath) Then
	'Make a temporary copy of the excel file
	fso.CopyFile strExcelFileNamePath, strExcelCopyFileNamePath, True
	
	'Open Excel and disable alerts
	Set oXL =  CreateObject("Excel.Application")
	'Wait one second for excel to properly open
	Delay 1
	oXL.DisplayAlerts = False
	'Open workbook and activate it
	Set oXLBook = oXL.Workbooks.Open(strExcelCopyFileNamePath, False, True)
	oXLBook.Activate
	'Choose desired worksheet
	Set oXLSheet = oXL.ActiveWorkbook.Worksheets(strSheetName)
	'Activate worksheet (same as clicking tab)
	oXLSheet.Activate
	'Retrieve value from worksheet
	strExcelOutputValue = oXLSheet.Cells(intExcelRow, intExcelColumn).Value
	
	'Handover value to EVOware
	Evoware.SetStringVariable "VBScript_GetExcelSheetCellValueString_String", cstr(strExcelOutputValue)
	
	'Quit workbook
	oXLBook.Close False
	Set oXLSheet = Nothing
	Set oXLBook = Nothing
	
	'Close excel instance
	oXL.Quit
	Set oXL = Nothing
	
	'Erase temporary excel file
	fso.DeleteFile strExcelCopyFileNamePath
Else
End If

'Housekeeping
Set strExcelFileNamePath = Nothing
Set intExcelRow = Nothing
Set intExcelColumn = Nothing
Set strExcelOutputValue = Nothing
Set strSheetName = Nothing
Set strExcelCopyFileNamePath = Nothing
Set fso = Nothing


Sub Delay(seconds) 'Submethod for delay without using Wscript
	Dim wshShell, strCmd
	Set wshShell = CreateObject( "WScript.Shell" )
	strCmd = wshShell.ExpandEnvironmentStrings( "%COMSPEC% /C (TIMEOUT.EXE /T " & seconds & " /NOBREAK)" )
	wshShell.Run strCmd, 0, 1
	Set wshShell = Nothing
End Sub