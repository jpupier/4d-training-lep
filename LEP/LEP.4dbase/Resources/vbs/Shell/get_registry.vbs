Set objShell 			= WScript.CreateObject ("WScript.Shell")
Set WshSysEnv 		= objShell.Environment("PROCESS")
WScript.StdOut.Write  	objShell.RegRead (WshSysEnv("SHELL_REGISTRY_KEY")) 
Set objShell	 		= Nothing