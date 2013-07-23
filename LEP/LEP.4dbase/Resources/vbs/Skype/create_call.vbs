Function GETENV(variableName)
	
	Set objWshShell 		= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv 		= objWshShell.Environment("PROCESS")
	GETENV 				= WshSysEnv(variableName)
	Set objWshShell 		= Nothing

end Function

Set objSkype			= WScript.CreateObject("Skype4COM.Skype", "4D")

If Not objSkype.Client.IsRunning Then objSkype.Client.Start() End If
If objSkype.Convert.TextToUserStatus("OFFLINE") = objSkype.CurrentUserStatus Then objSkype.ChangeUserStatus(objSkype.Convert.TextToUserStatus("OFFLINE")) End If

theTarget = GETENV("CALL_TARGET")

Set theCall = objSkype.PlaceCall(theTarget)

WScript.StdOut.Write theCall.Id