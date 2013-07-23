Function GETENV(variableName)
	
	Set objWshShell 		= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv 		= objWshShell.Environment("PROCESS")
	GETENV 				= WshSysEnv(variableName)
	Set objWshShell 		= Nothing

end Function

Set objFileSystemObject		= CreateObject("Scripting.FileSystemObject")

objFileSystemObject.MoveFile GETENV("FILE_SOURCE"), GETENV("FOLDER_DESTINATION")