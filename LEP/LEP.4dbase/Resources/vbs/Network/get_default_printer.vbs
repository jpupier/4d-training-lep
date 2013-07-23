Function GETENV(variableName)
	
	Set objWshShell 	= WScript.CreateObject("WScript.Shell")
	Set WshSysEnv 	= objWshShell.Environment("PROCESS")
	GETENV 			= WshSysEnv(variableName)
	Set objWshShell 	= Nothing

end Function

Set WshLocator = CreateObject("WbemScripting.SWbemLocator")
if Err = 0 then
      Set WshService = WshLocator.ConnectServer("","","","")
      if Err = 0 then
            WshService.Security_.impersonationlevel=3
            WshService.Security_.Privileges.AddAsString"SeLoadDriverPrivilege"
            Set WshEnum  = WshService.ExecQuery("select DeviceID from Win32_Printer where default=True")
            if Err.Number=0 then
                  for each WshPrinter in WshEnum
                        WScript.StdOut.Write WshPrinter.DeviceID
                        next
            end if
      end if
end if