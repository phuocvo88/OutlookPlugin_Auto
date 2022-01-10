if WinWaitActive("Windows Task Manager","",2) = 0 Then
	ConsoleWrite(WinWaitActive("Windows Task Manager","",2) & " no window " & @CRLF)
Else
ConsoleWrite(WinWaitActive("[ACTIVE]") & @CRLF)

EndIf