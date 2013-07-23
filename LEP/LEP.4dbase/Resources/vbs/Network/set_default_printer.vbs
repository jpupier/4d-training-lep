Function GETENV(variableName)
	
	Set objWshShell 	= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv 	= objWshShell.Environment("PROCESS")
	GETENV 			= WshSysEnv(variableName)
	Set objWshShell 	= Nothing

end Function

Set objNetwork = CreateObject("WScript.Network")
objNetwork.SetDefaultPrinter GETENV("NETWORK_PRINTER_NAME")