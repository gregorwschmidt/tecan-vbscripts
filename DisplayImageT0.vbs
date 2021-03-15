'Read path and title of image from Evoware
strImagePath = cstr(Evoware.GetStringVariable("VBScript_DisplayImage_FileNamePath"))
strImageTitle = cstr(Evoware.GetStringVariable("VBScript_DisplayImage_Title"))
strPopupMessage = cstr(Evoware.GetStringVariable("VBScript_DisplayImage_PopupMessage"))


'Extract width and height of image
dim iWidth, iHeight
dim oFs, oImg
Set oFs= CreateObject("Scripting.FileSystemObject")
Set oImg = loadpicture(strImagePath)
iWidth = round(oImg.width / 26.4583)
iHeight = round(oImg.height / 26.4583)
Set oImg = Nothing
Set oFs = Nothing

'Create internet explorer window showing the image
Set objExplorer = CreateObject("InternetExplorer.Application")

While objExplorer.Busy Or objExplorer.ReadyState <> READYSTATE_COMPLETE
    Delay 1
Wend


With objExplorer
    .Navigate "about:blank"
    .ToolBar = 0
    .StatusBar = 0
    .Left = 100
    .Top = 100
    .Width = iWidth+40
    .Height = iHeight+60
    .Visible = 1
End With

objExplorer.Document.Title = strImageTitle
objExplorer.Document.Head = "<meta http-equiv=x-ua-compatible content=IE=10>"
objExplorer.Document.Body.InnerHTML = "<img src='" & strImagePath & "'>"

'Bring IE window to the front
Set Processes = GetObject("winmgmts:").InstancesOf("Win32_Process")

intProcessId = ""
For Each Process In Processes
    If StrComp(Process.Name, "iexplore.exe", vbTextCompare) = 0 Then
        intProcessId = Process.ProcessId
        Exit For
    End If
Next

If Len(intProcessId) > 0 Then
    With CreateObject("WScript.Shell")
        .AppActivate intProcessId
   End With
End If

'Create popup message which has to be clicked
Set objShell = CreateObject("WScript.Shell")
	objShell.Popup strPopupMessage
Set objShell = Nothing


'Kill internet explorer
Set objExplorer = Nothing

strComputer = "."
strProcessToKill = "iexplore.exe" 

Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" _ 
	& strComputer & "\root\cimv2") 

Set colProcess = objWMIService.ExecQuery _
	("Select * from Win32_Process Where Name = '" & strProcessToKill & "'")

count = 0
For Each objProcess in colProcess
	objProcess.Terminate()
	count = count + 1
Next 

Sub Delay(seconds) 'Submethod for delay without using Wscript
	Dim wshShell, strCmd
	Set wshShell = CreateObject( "WScript.Shell" )
	strCmd = wshShell.ExpandEnvironmentStrings( "%COMSPEC% /C (TIMEOUT.EXE /T " & seconds & " /NOBREAK)" )
	wshShell.Run strCmd, 0, 1
	Set wshShell = Nothing
End Sub
