Function GETENV(variableName)
	
	Set objWshShell 		= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv			= objWshShell.Environment("PROCESS")
	GETENV					= WshSysEnv(variableName)
	Set objWshShell 		= Nothing

end Function

Function GETAPP(applicationName)

	On Error Resume Next
	Set GETAPP		 = GetObject(, applicationName)
	If Err.Number <> 0 Then
		Set GETAPP 	= CreateObject(applicationName) 
	End If
	On Error GoTo 0

end Function

Set objOutlook 				= GETAPP("Outlook.Application")
Set MAPI					= objOutlook.GetNameSpace("MAPI")
Set theDraftsFolder			= MAPI.GetDefaultFolder(16)
theDraftsFolder.Display

Set newMail					= objOutlook.CreateItem(0)

newMail.Subject				= GETENV("MESSAGE_SUBJECT")
newMail.Body				= GETENV("MESSAGE_CONTENT")

Set theRecipient			= newMail.Recipients.Add(GETENV("RECIPIENT_ADDRESS"))
theRecipient.Type			= 1

Set theRecipientCC			= newMail.Recipients.Add(GETENV("RECIPIENT_ADDRESS_CC"))
theRecipientCC.Type			= 2

Set theRecipientBCC			= newMail.Recipients.Add(GETENV("RECIPIENT_ADDRESS_BCC"))
theRecipientBCC.Type		= 3

Set theAttachment			= newMail.Attachments.Add(GETENV("ATTACHMENT_PATH"))

newMail.Display

Set objOutlook = Nothing