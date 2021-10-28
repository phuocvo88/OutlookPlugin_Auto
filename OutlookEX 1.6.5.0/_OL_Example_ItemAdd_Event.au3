#include <OutlookEX.au3>

; *****************************************************************************
; Example Script
; Handle Outlook Folder event when a new item arrives in a folder.
; This script checks the Sent Items folder to display a message when a mail
; has been sent (= a copy of the sent mail is stored in this folder).
; This script loops until Shift-Alt-E is pressed to exit.
; *****************************************************************************
HotKeySet("+!e", "_Exit") ;Shift-Alt-E to Exit the script
MsgBox($MB_IconInformation, "OutlookEX UDF Example Script", "Hotkey to exit the script: 'Shift-Alt-E'!")

Global $oOL = ObjCreate("Outlook.Application")
Global $oItems = $oOL.GetNamespace("MAPI").GetDefaultFolder($olFolderSentMail).Items
ObjEvent($oItems, "oItems_")

While 1
	Sleep(10)
WEnd

; ItemAdd event - https://docs.microsoft.com/en-us/office/vba/api/outlook.items.itemadd
Func oItems_ItemAdd($oOL_Item)
	MsgBox($MB_ICONINFORMATION, "OutlookEX UDF Example Script", "Mail has been sent!" & @CRLF & @CRLF & _
			"Subject: " & $oOL_Item.Subject)
EndFunc   ;==>oItems_ItemAdd

Func _Exit()
	Exit
EndFunc   ;==>_Exit
