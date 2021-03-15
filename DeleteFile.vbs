'This script deletes a file
'
'Inputs:
'	Evoware variable "VBScript_DeleteFile_FilePath": Full path to file
'
'Gregor Schmidt (gregor.schmidt@bsse.ethz.ch), 07/01/2021

'Housekeeping
Option Explicit
Dim strFilePath
Dim fso

'Read variables from evoware
strFilePath = CStr(Evoware.GetStringVariable("VBScript_DeleteFile_FilePath"))

'Create file system object
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(strFilePath) Then
	'Delete file
	fso.DeleteFile strFilePath
End If


'Housekeeping
Set fso = Nothing
Set strFilePath = Nothing  