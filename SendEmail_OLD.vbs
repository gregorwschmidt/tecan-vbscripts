 Dim ToAddress
 Dim MessageSubject
 Dim MessageBody 
 Dim ol, ns, newMail
 
 ToAddress = Evoware.GetStringVariable("VBScript_SendMail_Recipient")
 MessageSubject = Evoware.GetStringVariable("VBScript_SendMail_Subject")
 MessageBody = Evoware.GetStringVariable("VBScript_SendMail_Message")
  
 Set ol = CreateObject("Outlook.Application")
 Set ns = ol.getNamespace("MAPI")
 ns.logon "","",false,false
 Set newMail = ol.CreateItem(olMailItem)
 newMail.Subject = MessageSubject
 newMail.Body = MessageBody & vbCrLf
 newMail.Recipients.Add("laf_staff@bsse.ethz.ch")
 newMail.Recipients.Add(ToAddress)
 newMail.Send  
 Set ol = Nothing
 
 Dim WshShell 
 Dim ObjOL 'As Outlook.Application
 
 'Start Outlook to make sure mail is not stuck in outbox
 set WshShell = CreateObject("WScript.Shell")
 WshShell.Run "outlook"
 Delay 2
 WshShell.AppActivate "Outlook"
 Delay 1
 WshShell.SendKeys "{F9}" 'Shortcut for mail sending
 Delay 3

 Set ObjOL = CreateObject("Outlook.Application")
 If ObjOL Is Nothing Then
 'Outlook is not running - all good.
 Else
 'Outlook is running - turn it off.
 ObjOL.Session.Logoff
 ObjOL.Quit
 End If
 Set ObjOL = Nothing
 
 ' Submethod for delay without using Wscript
Sub Delay( seconds )
	Dim wshShell, strCmd
	Set wshShell = CreateObject( "WScript.Shell" )
	strCmd = wshShell.ExpandEnvironmentStrings( "%COMSPEC% /C (TIMEOUT.EXE /T " & seconds & " /NOBREAK)" )
	wshShell.Run strCmd, 0, 1
	Set wshShell = Nothing
End Sub