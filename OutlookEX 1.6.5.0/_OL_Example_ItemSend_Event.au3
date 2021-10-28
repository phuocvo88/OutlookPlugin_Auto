#include <OutlookEX.au3>

; *****************************************************************************
; Example Script
; Handle Outlook event when a new item is being sent.
; This example script checks for mail items and changes the subject before the mail is
; being sent.
; This script loops until Shift-Alt-E is pressed to exit.
; *****************************************************************************
HotKeySet("+!e", "_Exit") ;Shift-Alt-E to Exit the script
MsgBox(64, "OutlookEX UDF Example Script", "Hotkey to exit the script: 'Shift-Alt-E'!")

Global $oOL = _OL_Open()
Global $oEvent = ObjEvent($oOL, "_oItems_")

While 1
	Sleep(50)
WEnd

; ItemSend event - https://docs.microsoft.com/en-us/office/vba/api/outlook.application.itemsend
Volatile Func _oItems_ItemSend($oItem, $bCancel)
	#forceref $bCancel
	$bCancel = False ; If you do not want to send the item then set $bCancel to True
	If $oItem.Class = $olMail Then
		$oItem.Subject = "Modified by AutoIt script: " & $oItem.Subject
		$oItem.Save()
	EndIf
EndFunc   ;==>oItems_ItemSend

Func _Exit()
	Exit
EndFunc   ;==>_Exit
