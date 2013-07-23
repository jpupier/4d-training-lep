Function GETENV(variableName)
	
	Set objWshShell 		= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv 		= objWshShell.Environment("PROCESS")
	GETENV 				= WshSysEnv(variableName)
	Set objWshShell 		= Nothing

end Function

Set objShell 				= CreateObject("Shell.Application")
theFolderPath				= GETENV("FOLDER_PATH")
Set theFolder 				= objShell.NameSpace(theFolderPath)
Set theTrash 				= objShell.NameSpace(10)

theTrash.MoveHere theFolder.Self, 0

Do
	If objShell.NameSpace(theFolderPath) is Nothing Then
		Exit Do
	end if
	WScript.Sleep 500
Loop