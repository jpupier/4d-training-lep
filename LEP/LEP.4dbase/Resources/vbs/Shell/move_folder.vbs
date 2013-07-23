Function GETENV(variableName)
	
	Set objWshShell 		= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv 		= objWshShell.Environment("PROCESS")
	GETENV 				= WshSysEnv(variableName)
	Set objWshShell 		= Nothing

end Function

Set objShell 				= CreateObject("Shell.Application")
Set theSource 				= objShell.NameSpace(GETENV("FOLDER_SOURCE"))
Set theDestination 			= objShell.NameSpace(GETENV("FOLDER_DESTINATION"))

theDestination.MoveHere theSource.Self, 16