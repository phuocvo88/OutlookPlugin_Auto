#include <OutlookEX.au3>

; *****************************************************************************
; Handle Outlook events when a built-in property of an item is changed.
; Select at least one item in the active Explorer.
; You get a message written to the Console for every item when the property "Unread" changes
; This script loops until Shift-Alt-E is pressed to exit.
; *****************************************************************************
HotKeySet("+!e", "_Exit") 								; Shift-Alt-E to exit the script
MsgBox(64, "OutlookEX UDF Example Script", "Hotkey to exit the script: 'Shift-Alt-E'!")

Global $oItem, $oEvent
Global Const $olViewList = 0 							; The selection is in a list of items in an explorer
Global $oOL = _OL_Open()							    ; Connect to Outlook
Global $oExplorer = $oOL.ActiveExplorer()				; Get the active Explorer object
ObjEvent($oExplorer, "oExplorer_")						; Define an event for a selection change
$oItem = $oExplorer.Selection(1) 						; Get first item of the selection
$oEvent = ObjEvent($oItem, "oItem_") 					; Define an event for the first item of the selection
While 1
	Sleep(10)
WEnd

Func oExplorer_SelectionChange()
	If $oExplorer.Selection.Location = $olViewList Then	; Selection changed in the view part of the Explorer
		$oItem = $oExplorer.Selection(1)				; Get first item of the new selection
		$oEvent.Stop									; Stop receiving events for the last selection
		$oEvent = ObjEvent($oItem, "oItem_") 			; Define an event for the first item of the new selection
	EndIf
EndFunc   ;==>oExplorer_SelectionChange

Func oItem_PropertyChange($sChangedProperty)
	If $sChangedProperty = "UnRead" Then
		For $oSelected In $oExplorer.Selection
			ConsoleWrite("Subject: " & $oSelected.Subject & ", Value of property '" & $sChangedProperty & "' before change: " & $oSelected.Unread & @CRLF)
		Next
	EndIf
EndFunc   ;==>oMailSelected_PropertyChange

Func _Exit()
	Exit
EndFunc   ;==>_Exit
