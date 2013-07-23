Function GETENV(variableName)
	
	Set objWshShell 	= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv 	= objWshShell.Environment("PROCESS")
	GETENV 			= WshSysEnv(variableName)
	Set objWshShell 	= Nothing

end Function

Set objSystem = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")

for each System in objSystem
	WScript.StdOut.Write System.Version
next