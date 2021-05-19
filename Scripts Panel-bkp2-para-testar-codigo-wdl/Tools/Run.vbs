Set objShell = CreateObject( "WScript.Shell" )
objShell.Run "%comspec% /C " & arguments(0) , 0, True



