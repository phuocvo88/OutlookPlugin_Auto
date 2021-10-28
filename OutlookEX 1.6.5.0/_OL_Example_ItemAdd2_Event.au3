#include <OutlookEX.au3>
#include <MsgBoxConstants.au3>

; *****************************************************************************
; Example Script
; Handle Outlook ItemAdd event when a new mail item arrives in an Inbox.
; This example works for an Inbox of another user or a shared mailbox.
; This script loops until Shift-Alt-E is pressed to exit.
; *****************************************************************************
HotKeySet("+!e", "_Exit") ; Shift-Alt-E to Exit the script
Global $sMailbox = "mailbox@company.com"     										  ; <== replace with the mailbox you want to monitor
Global $sTitle = "OutlookEX UDF Example Script"
MsgBox($MB_IconInformation, $sTitle, "Hotkey to exit the script: 'Shift-Alt-E'!", 10) ; Wait 10 seconds, then continue

; Start or connect to a running Outlook instance
Global $oOL = _OL_Open()
If @error Then Exit MsgBox($MB_IconError, $sTitle, "Error when calling _OL_Open: @error=" & @error & ", @extended=" & @extended & @CRLF)

; Access the Mailbox
Global $aFolder = _OL_FolderAccess($oOL, $sMailbox, $olFolderInbox)
If @error Then Exit MsgBox($MB_IconError, $sTitle, "Error when calling _OL_FolderAccess: @error=" & @error & ", @extended=" & @extended & @CRLF)
_Arraydisplay($aFolder)
Exit

; Create a collection of the items in this mailbox
Global $oItems = $aFolder[1].Items
If @error Then Exit MsgBox($MB_IconError, $sTitle, "Error when accessing the folder items: @error=" & @error & ", @extended=" & @extended & @CRLF)

; Create the events for this collection. Outlook calls a function starting with "oOL_" for each event. For the ItemAdd event function oOL_ItemAdd will be called
Global $oTemp = ObjEvent($oItems, "oOL_")
If @error Then Exit MsgBox($MB_IconError, $sTitle, "Error when calling ObjEvent: @error=" & @error & ", @extended=" & @extended & @CRLF)
ConsoleWrite("OutlookEX UDF Example Script - waiting for new items to arrive!" & @CRLF)

While 1
	Sleep(10)
WEnd

; ItemAdd event - https://docs.microsoft.com/en-us/office/vba/api/outlook.items.itemadd
Func oOL_ItemAdd($oItem)
	ConsoleWrite("OutlookEX UDF Example Script - new item has arrived!" & @CRLF)
	ConsoleWrite( _
			"From:    " & $oItem.SenderName & @CRLF & _
			"Subject: " & $oItem.Subject & @CRLF & _
			"Class:   " & $oItem.Class & " (43=Mail, 53=MeetingRequest ...)" & @CRLF)
EndFunc   ;==>oOL_ItemAdd

Func _Exit()
	_OL_Close($oOL)
	Exit
EndFunc   ;==>_Exit
