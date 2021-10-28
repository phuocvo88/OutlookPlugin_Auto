#include <OutlookEX.au3>

; *****************************************************************************
; Connect to Outlook
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Start synchronization for all Send/Receive groups
; *****************************************************************************
_OL_Sync($oOutlook)
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_Sync Example Script", "Error synchronizing the specified Send/Receive groups. @error = " & @error & ", @extended: " & @extended)
MsgBox($MB_ICONINFORMATION, "OutlookEX UDF: _OL_Sync Example Script", "Successfully synced " & @extended & " Send/Receive groups!")

_OL_Close($oOutlook)