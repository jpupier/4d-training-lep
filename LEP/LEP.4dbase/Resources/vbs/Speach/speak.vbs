Function GETENV(variableName)
	
	Set objWshShell 		= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv 		= objWshShell.Environment("PROCESS")
	GETENV 				= WshSysEnv(variableName)
	Set objWshShell 		= Nothing

end Function

Set ObjSpVoice				= CreateObject("SAPI.SpVoice")
ObjSpVoice.Rate 			= CLng(GETENV("SPEAK_RATE"))
ObjSpVoice.Volume 			= CLng(GETENV("SPEAK_VOLUME"))

ObjSpVoice.Speak GETENV("SPEAK_MESSAGE")