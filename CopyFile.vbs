'This script copies a file to a destination folder
'
'Inputs:
'	Evoware variable "VBScript_CopyFile_SourceFilePath": Full path of source file
'	Evoware variable "VBScript_CopyFile_DestinationFolder": Full path of destination folder
'	Evoware variable "VBScript_CopyFile_DestinationFileName": New Name of file in destination folder. If this is left empty e.g. " ", the file will be copied and keep its old name.
'Outputs:
'	Evoware variable "VBScript_CopyFile_SourceFolder": Full path of source folder
'	Evoware variable "VBScript_CopyFile_DestinationFilePath": Full path of destination file

'Read variables from evoware
SourceFilePath = CStr(Evoware.GetStringVariable("VBScript_CopyFile_SourceFilePath"))
DestinationFolder = CStr(Evoware.GetStringVariable("VBScript_CopyFile_DestinationFolder"))
DestinationFileName = CStr(Evoware.GetStringVariable("VBScript_CopyFile_DestinationFileName"))

'General housekeeping
Dim fso

'Create file system object
Set fso = CreateObject("Scripting.FileSystemObject")

'Add a backslash if slash at end of destination folder is missing
If  Right(DestinationFolder,1) <>"\"Then
	DestinationFolder=DestinationFolder&"\"
End If

'Extract filename of source file
SourceFileName = Mid(SourceFilePath, InStrRev(SourceFilePath, "\") + 1)


If DestinationFileName = " " Then
	DestinationFilePath = DestinationFolder & SourceFileName
Else
	DestinationFilePath = DestinationFolder & DestinationFileName
End if


'Copy the file
fso.CopyFile SourceFilePath, DestinationFilePath, True

'Extract folder from source file path
SourceFolder =  fso.GetParentFolderName(SourceFilePath)&"\"

'Delete file system object
Set fso = Nothing

'Write destination file path to evoware
Evoware.SetStringVariable "VBScript_CopyFile_DestinationFilePath", DestinationFilePath

'Write source folder to evoware
Evoware.SetStringVariable "VBScript_CopyFile_SourceFolder", SourceFolder


