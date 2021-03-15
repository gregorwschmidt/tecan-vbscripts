'This script copies a file to a destination folder
'
'Inputs:
'	Evoware variable "VBScript_CopyEvowareMethodFiles_ProjectID": Project ID e.g. P038
'	Evoware variable "VBScript_CopyEvowareMethodFiles_DestinationFolder": Full path of destination folder

'Read variables from evoware
ProjectID = CStr(Evoware.GetStringVariable("VBScript_CopyEvowareMethodFiles_ProjectID"))
DestinationFolder = CStr(Evoware.GetStringVariable("VBScript_CopyEvowareMethodFiles_DestinationFolder"))

'General housekeeping
Dim fso

'Create file system object
Set fso = CreateObject("Scripting.FileSystemObject")

'Add a backslash if slash at end of destination folder is missing
If  Right(DestinationFolder,1) <>"\"Then
	DestinationFolder=DestinationFolder&"\"
End If

MethodsFolder = "C:\ProgramData\Tecan\EVOware\database\scripts\"
WorktableFolder = "C:\ProgramData\Tecan\EVOware\database\wt_templates\"
CustomLCPath =  "C:\ProgramData\Tecan\EVOware\database\CustomLCs.XML"
DefaultLCPath =  "C:\ProgramData\Tecan\EVOware\database\DefaultLCs.XML"


'Check script folder for method with the correct project ID and copy them
Set folder = fso.getfolder(MethodsFolder)
    For Each file in folder.Files
      If inStr(file.name, ProjectID)<>0 Then
		fso.CopyFile File, DestinationFolder, True
      End If
    Next

Set folder = nothing

'Check worktable folder for worktable with the correct project ID and copy them
Set folder = fso.getfolder(WorktableFolder)
    For Each file in folder.Files
      If inStr(file.name, ProjectID)<>0 Then
		fso.CopyFile File, DestinationFolder, True
      End If
    Next

Set folder = nothing

'Copy the liquid class files
If fso.FileExists (CustomLCPath) Then
fso.CopyFile CustomLCPath, DestinationFolder, True
End If
If fso.FileExists (DefaultLCPath) Then
fso.CopyFile DefaultLCPath, DestinationFolder, True
End If


'Delete file system object
Set fso = Nothing




