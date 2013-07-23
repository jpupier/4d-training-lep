Function GETENV(variableName)
	
	Set objShell 		= WScript.CreateObject("WScript.Shell")
	Set theVariable		= objShell.Environment("PROCESS")
	GETENV 				= theVariable(variableName)
	Set objShell 		= Nothing

end Function

Set objShell			= WScript.CreateObject("Shell.Application")

objShell.TileHorizontally
