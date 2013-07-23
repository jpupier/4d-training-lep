Function GETENV(variableName)
	
	Set objWshShell 		= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv 		= objWshShell.Environment("PROCESS")
	GETENV 				= WshSysEnv(variableName)
	Set objWshShell 		= Nothing

end Function

Set objEmail 				= CreateObject("CDO.Message")
objEmail.From				= GETENV("MAIL_FROM")
objEmail.To				= GETENV("MAIL_TO")
objEmail.Cc				= GETENV("MAIL_CC")
objEmail.Bcc				= GETENV("MAIL_BCC")
objEmail.Subject			= GETENV("MAIL_SUBJECT")
objEmail.HTMLBody			= WScript.StdIn.ReadAll

attachmentCount = 1

Do While 1
	attachment = GETENV("MAIL_ATTACHMENT_" & attachmentCount)
	If attachment <> "" Then
		objEmail.AddAttachment attachment
		attachmentCount = attachmentCount + 1
	Else
		Exit Do
	End If
Loop

objEmail.AutoGenerateTextBody = True

objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") 			= GETENV("MAIL_CONFIGURAION_MODE")
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") 			= GETENV("MAIL_CONFIGURAION_SERVER")
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") 	= GETENV("MAIL_CONFIGURAION_TIMEOUT")
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") 		= GETENV("MAIL_CONFIGURAION_PORT")
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") 		= GETENV("MAIL_CONFIGURAION_AUTH")
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") 			= GETENV("MAIL_CONFIGURAION_USER")
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") 			= GETENV("MAIL_CONFIGURAION_PASSWORD")
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") 			= GETENV("MAIL_CONFIGURAION_USESSL")

objEmail.Configuration.Fields.Update
WScript.StdOut.Write objEmail.Send

Set objEmail 		= Nothing