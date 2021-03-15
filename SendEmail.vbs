 Dim ToAddress
 Dim MessageSubject
 Dim MessageBody
 Dim CommandLineString
 
 ToAddress = Evoware.GetStringVariable("VBScript_SendMail_Recipient")
 MessageSubject = Evoware.GetStringVariable("VBScript_SendMail_Subject")
 MessageBody = Evoware.GetStringVariable("VBScript_SendMail_Message")
 
 'create command line string to call later
 CommandLineString = """C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe"" -compose ""to='laf_staff@bsse.ethz.ch"
 'add user email address to command string
 CommandLineString = CommandLineString & "," & ToAddress & "'"
 'add email subject to command string
 CommandLineString = CommandLineString & ",subject='" & MessageSubject & "'"
 'add body to email
 CommandLineString = CommandLineString & ",body='" & MessageBody & "'"""
 
 
 'Create shell object and run command line created earlier
 Set WshShell = CreateObject("WScript.Shell")
 WshShell.run CommandLineString
 Delay 5
 'Send message using shortcut keys
 WshShell.SendKeys "^{ENTER}" 'Shortcut for mail sending
 Delay 1
 
 'Housekeeping
 Set WshShell = Nothing
 Set ToAddress = Nothing
 Set MessageSubject = Nothing
 Set MessageBody = Nothing
 Set CommandLineString = Nothing
 
 'Submethod for delay without using Wscript
Sub Delay( seconds )
	Dim wshShellDelay, strCmd
	Set wshShellDelay = CreateObject( "WScript.Shell" )
	strCmd = wshShellDelay.ExpandEnvironmentStrings( "%COMSPEC% /C (TIMEOUT.EXE /T " & seconds & " /NOBREAK)" )
	wshShellDelay.Run strCmd, 0, 1
	Set wshShellDelay = Nothing
End Sub