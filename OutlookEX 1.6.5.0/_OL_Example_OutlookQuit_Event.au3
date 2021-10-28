#include <OutlookEX.au3>

; *****************************************************************************
; Example Script
; Handle Outlook Quit event when Outlook is being closed.
; This script loops until Shift-Alt-E is pressed to exit.
; *****************************************************************************
HotKeySet("+!e", "_Exit") ;Shift-Alt-E to Exit the script
MsgBox(64, "OutlookEX UDF Example Script", "Hotkey to exit the script: 'Shift-Alt-E'!")



Global $oOL = _OL_Open()

$iPID = ProcessExists("OUTLOOK.EXE")
ConsoleWrite("PID of Outlook = " & $iPID)
$oItem = _OL_ItemCreate
$oItem.Display
_OL_ItemAttachmentAdd($oOL, @ScriptDir & "\The_Outlook.jpg")

Sleep(10000)

While 1
	Sleep(10)
WEnd

; Quit event - https://docs.microsoft.com/en-us/office/vba/api/outlook.application.quit(even)
Func oOApp_Quit()
	MsgBox($MB_ICONINFORMATION, "OutlookEX UDF Example Script", "Outlook is being closed!" & @CRLF & "Goodby ")
EndFunc   ;==>oOApp_NewMailEx

Func _Exit()
	Exit
EndFunc   ;==>_Exit