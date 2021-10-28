#include <OutlookEX.au3>

; *****************************************************************************
; Example Script
; Handle Outlook NewmailEX event when one/multiple new mail(s) arrive(s) in your Inbox.
; This script loops until Shift-Alt-E is pressed to exit.
; *****************************************************************************
HotKeySet("+!e", "_Exit") ;Shift-Alt-E to Exit the script
MsgBox(64, "OutlookEX UDF Example Script", "Hotkey to exit the script: 'Shift-Alt-E'!")

Global $oOL = _OL_Open()
Global $oEvent = ObjEvent($oOL, "oOApp_")

While 1
	Sleep(10)
WEnd

; NewMailEx event - https://docs.microsoft.com/en-us/office/vba/api/outlook.application.newmailex
Func oOApp_NewMailEx($sEntryIDs)

	Local $iItemCount, $oItem, $aEntryIDs = StringSplit($sEntryIDs, ",", $STR_NOCOUNT)
	$iItemCount = UBound($aEntryIDs)
	ConsoleWrite("OutlookEX UDF Example Script - " & ($iItemCount = 1 ? "new item has" : "new items have") & " arrived!" & @CRLF & @CRLF)
	For $i = 0 To $iItemCount - 1
		$oItem = $oOL.Session.GetItemFromID($aEntryIDs[$i], Default)
		ConsoleWrite( _
				"From:    " & $oItem.SenderName & @CRLF & _
				"Subject: " & $oItem.Subject & @CRLF & _
				"Class:   " & $oItem.Class & " (43=Mail, 53=MeetingRequest ...)" & @CRLF)
	Next
EndFunc   ;==>oOApp_NewMailEx

Func _Exit()
	Exit
EndFunc   ;==>_Exit
