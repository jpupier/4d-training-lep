Function GETENV(variableName)
	
	Set objWshShell 		= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv 		= objWshShell.Environment("PROCESS")
	GETENV 				= WshSysEnv(variableName)
	Set objWshShell 		= Nothing

end Function

Set objSkype			= WScript.CreateObject("Skype4COM.Skype", "4D")

If Not objSkype.Client.IsRunning Then objSkype.Client.Start() End If
If objSkype.Convert.TextToUserStatus("OFFLINE") = objSkype.CurrentUserStatus Then objSkype.ChangeUserStatus(objSkype.Convert.TextToUserStatus("OFFLINE")) End If

theSMSID = CLng(GETENV("SMS_ID"))

For i = 1 To objSkype.Smss.Count
   If objSkype.Smss.Item(i).Id = theSMSID Then
	WScript.StdOut.Write objSkype.Smss.Item(i).Body   
       Exit For
   End If
Next