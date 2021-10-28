#include <OutlookEX.au3>

; *****************************************************************************
; Connect to Outlook
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; List all Outlook profiles
; *****************************************************************************
Global $aResult = _OL_ProfileGet($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ProfileGet Example Script", "Error getting list of profiles. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aResult, "OutlookEX UDF: _OL_ProfileGet Example Script", "", 0, "|", "Profile name|Default profile?")

_OL_Close($oOutlook)