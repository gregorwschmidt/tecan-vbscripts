Option Explicit

Dim Path

'Get path variable from evoware and create path
Path = CStr(Evoware.GetStringVariable("VBScript_CreateFolder_Path"))
CreateFolder(Path)


'Function to create path
Function CreateFolder(NewPath)

Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")
'Create new folder unless it already exists
If Not fso.FolderExists(NewPath) then
fso.CreateFolder(NewPath)
End If

Set fso = Nothing

End Function

