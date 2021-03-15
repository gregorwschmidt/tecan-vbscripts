'This script copies a file to a destination folder
'
'Inputs:
'	Evoware variable "VBScript_CopyFileSafe_SourceFilePath": Full path of source file
'	Evoware variable "VBScript_CopyFileSafe_DestinationFolder": Full path of destination folder
'	Evoware variable "VBScript_CopyFileSafe_DestinationFileName": New Name of file in destination folder. If this is left empty e.g. " ", the file will be copied and keep its old name.
'Outputs:
'	Evoware variable "VBScript_CopyFileSafe_SourceFolder": Full path of source folder
'	Evoware variable "VBScript_CopyFileSafe_DestinationFilePath": Full path of destination file
'
'Gregor Schmidt (gregor.schmidt@bsse.ethz.ch), 09/06/2020

'Housekeeping
Option Explicit
Dim strSourceFilePath,strSourceFolder, strSourceFileName, strDestinationFilePath, strDestinationFolder, strDestinationFileName
Dim fso, intReturnCode

'Read variables from evoware
strSourceFilePath = CStr(Evoware.GetStringVariable("VBScript_CopyFileSafe_SourceFilePath"))
strDestinationFolder = CStr(Evoware.GetStringVariable("VBScript_CopyFileSafe_DestinationFolder"))
strDestinationFileName = CStr(Evoware.GetStringVariable("VBScript_CopyFileSafe_DestinationFileName"))

'Create file system object
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(strSourceFilePath) Then
	'Add a backslash if slash at end of destination folder in case it is missing
	If  Right(strDestinationFolder,1) <>"\"Then
		strDestinationFolder=strDestinationFolder&"\"
	End If
	
	'Extract filename of source file
	strSourceFileName = Mid(strSourceFilePath, InStrRev(strSourceFilePath, "\") + 1)
	
	'Build destination file path
	If strDestinationFileName = " " Then
		strDestinationFilePath = strDestinationFolder & strSourceFileName
	Else
		strDestinationFilePath = strDestinationFolder & strDestinationFileName
	End if

	'Copy the file
	fso.CopyFile strSourceFilePath, strDestinationFilePath, True
	
	
	'Extract folder from source file path
	strSourceFolder =  fso.GetParentFolderName(strSourceFilePath)&"\"
	
	'Write source folder to evoware
	Evoware.SetStringVariable "VBScript_CopyFileSafe_SourceFolder", strSourceFolder
	
	'Write destination file path to evoware
	Evoware.SetStringVariable "VBScript_CopyFileSafe_DestinationFilePath", strDestinationFilePath
	
	intReturnCode = 0
Else
	intReturnCode = 1
End If

'Return code back to EVOware to indicate if copying was successful
Evoware.SetDoubleVariable "VBScript_CopyFileSafe_ReturnCode", intReturnCode

'Housekeeping
Set fso = Nothing
Set intReturnCode = Nothing

Set strSourceFilePath = Nothing
Set strSourceFolder = Nothing
Set strSourceFileName = Nothing

Set strDestinationFilePath = Nothing
Set strDestinationFolder = Nothing
Set strDestinationFileName = Nothing



























