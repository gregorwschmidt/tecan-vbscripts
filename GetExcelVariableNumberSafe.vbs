Option Explicit
Dim strExcelFileNamePath, strExcelVariableName
Dim oXL, oXLBook, oXLSheet, objShell, floExcelOutputValue, fso, strExcelFilePath, Temp, strFileExtension, strExcelCopyFileNamePath, dateTime, intReturnCode

' Read values from evoware
strExcelFileNamePath = cstr(Evoware.GetStringVariable("VBScript_GetExcelVariableNumberSafe_FileNamePath"))
strExcelVariableName = cstr(Evoware.GetStringVariable("VBScript_GetExcelVariableNumberSafe_VariableName"))


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
	
	
	
	'Retrieve value	
	floExcelOutputValue = cdbl(oXLBook.Worksheets(1).Range(strExcelVariableName).Value)
	
	'Handover value to EVOware
	Evoware.SetDoubleVariable strExcelVariableName, floExcelOutputValue
		
	'Quit workbook
	oXLBook.Close False
	Set oXLSheet = Nothing
	Set oXLBook = Nothing
	
	'Close excel instance
	oXL.Quit
	Set oXL = Nothing
	
	'Erase temporary excel file
	fso.DeleteFile strExcelCopyFileNamePath
	
	intReturnCode = 0
Else
	intReturnCode = 1
End If

Evoware.SetDoubleVariable "VBScript_GetExcelVariableNumberSafe_ReturnCode", intReturnCode

'Housekeeping
Set intReturnCode = Nothing
Set strExcelFileNamePath = Nothing
Set strExcelVariableName = Nothing
Set floExcelOutputValue = Nothing
Set strExcelCopyFileNamePath = Nothing
Set fso = Nothing


Sub Delay(seconds) 'Submethod for delay without using Wscript
	Dim wshShell, strCmd
	Set wshShell = CreateObject( "WScript.Shell" )
	strCmd = wshShell.ExpandEnvironmentStrings( "%COMSPEC% /C (TIMEOUT.EXE /T " & seconds & " /NOBREAK)" )
	wshShell.Run strCmd, 0, 1
	Set wshShell = Nothing
End Sub