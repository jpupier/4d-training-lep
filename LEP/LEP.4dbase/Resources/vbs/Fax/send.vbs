Function GETENV(variableName)
	
	Set objWshShell 	= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv 	= objWshShell.Environment("PROCESS")
	GETENV 			= WshSysEnv(variableName)
	Set objWshShell 	= Nothing

end Function

Set objFaxDocument = WScript.CreateObject ("FaxComEx.FaxDocument")

objFaxDocument.Body 			= GETENV("FAXDOCUMENT_BODY")
objFaxDocument.DocumentName 	= GETENV("FAXDOCUMENT_DOCUMENTNAME") 
objFaxDocument.Recipients.Add(GETENV("FAXDOCUMENT_RECIPIENT")) 
objFaxDocument.Sender.Name 		= GETENV("FAXDOCUMENT_SENDER_NAME")
objFaxDocument.Sender.FaxNumber 	= GETENV("FAXDOCUMENT_SENDER_FAX")
objFaxDocument.Sender.Email 		= GETENV("FAXDOCUMENT_SENDER_EMAIL")
objFaxDocument.Sender.Company	= GETENV("FAXDOCUMENT_SENDER_COMPANY")

Set objFaxServer = WScript.CreateObject("FaxComEx.FaxServer")

FaxServer.Connect GETENV("FAXSERVER")

JobID = FaxDoc.ConnectedSubmit(FaxServer) 

WScript.StdOut.Write JobID