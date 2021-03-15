Option Explicit
Dim strExcelFileNamePath, oXL, oXLBook, objShell, strMacroName
Dim fso, strExcelFilePath, Temp, strFileExtension, strExcelCopyFileNamePath, dateTime, intReturnCode

' Read values from evoware
strExcelFileNamePath = cstr(Evoware.GetStringVariable("VBScript_RunMacroExcelFile_FileNamePath"))
strMacroName = cstr(Evoware.GetStringVariable("VBScript_RunMacroExcelFile_MacroName"))

'Create file system object
Set fso = CreateObject("Scripting.FileSystemObject")

'Build file path for temporary excel file 
strExcelFilePath =  fso.GetParentFolderName(strExcelFileNamePath)&"\"
Temp = split(strExcelFileNamePath, ".")
strFileExtension = Temp(1)
dateTime = Year(now) & "-" & right("00" & Month(now),2) & "-" & right("00" & Day(now),2) & "_" & right("00" & Hour(now),2) & "-" & right("00" & Minute(now),2) & "-" & right("00" & Second(now),2)
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
	'Hide excel
	oXl.Application.Visible = False
	'Hide any excel alerts
	oXL.DisplayAlerts = False
	'Open copied excel file and activate it
	Set oXLBook = oXL.Workbooks.Open(strExcelCopyFileNamePath, False, True)
	oXLBook.Activate
	'Run macro
	oXL.Run(strMacroName)
	'Close excel file
	oXLBook.Close False
	Set oXLBook = Nothing
	'Close excel instance
	oXL.Quit
	Set oXL = Nothing
	'Wait for excel to close, such that we can erase file in the next step
	Delay 1
	'Erase temporary excel file
	fso.DeleteFile strExcelCopyFileNamePath
	
	intReturnCode = 0
Else
	intReturnCode = 1
End If

Evoware.SetDoubleVariable "VBScript_RunMacroExcelFile_ReturnCode", intReturnCode

'Housekeeping
Set intReturnCode = Nothing
Set strExcelFileNamePath = Nothing
Set strMacroName = Nothing
Set strExcelCopyFileNamePath = Nothing
Set fso = Nothing


Sub Delay(seconds) 'Submethod for delay without using Wscript
	Dim wshShell, strCmd
	Set wshShell = CreateObject( "WScript.Shell" )
	strCmd = wshShell.ExpandEnvironmentStrings( "%COMSPEC% /C (TIMEOUT.EXE /T " & seconds & " /NOBREAK)" )
	wshShell.Run strCmd, 0, 1
	Set wshShell = Nothing
End Sub