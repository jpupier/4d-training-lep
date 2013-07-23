Function GETENV(variableName)
	
	Set objWshShell 		= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv 		= objWshShell.Environment("PROCESS")
	GETENV 				= WshSysEnv(variableName)
	Set objWshShell 		= Nothing

end Function

Set objShell 				= CreateObject("Shell.Application")
theFolderPath				= GETENV("FOLDER_PATH")
theFileName				= GETENV("FILE_NAME")
Set theFolder				= objShell.NameSpace(theFolderPath)
Set theFile					= theFolder.ParseName(theFileName)
Set theTrash 				= objShell.NameSpace(10)

theTrash.MoveHere theFile, 0

Do
	If theFolder.ParseName(theFileName) is Nothing Then
		Exit Do
	end if
	WScript.Sleep 100
Loop