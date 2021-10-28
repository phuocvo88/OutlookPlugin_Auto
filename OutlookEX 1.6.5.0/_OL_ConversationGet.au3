#include <OutlookEX.au3>

If MsgBox(0, "_OL_ConversionGet", "Please select an email item!") <> $IDOK Then Exit

; *****************************************************************************
; Connect to Outlook
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Retrieve all items with the same conversation ID as the selected item
; *****************************************************************************
Global $oItem = $oOutlook.ActiveExplorer().Selection(1)
; If $oItem.Class <> $olMail Then Exit MsgBox($MB_TOPMOST, "_OL_ConversionGet", "Selected item has to be an email!")
Global $aResult = _OL_ConversationGet($oOutlook, $oItem)
If @error Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF", "Error retrieving the conversation for the specified item. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aResult, Default, Default, Default, Default, "EntryID|Subject|CreationTime|LastModificationTime|MessageClass")
