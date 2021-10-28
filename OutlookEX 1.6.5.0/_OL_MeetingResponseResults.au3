; #include <OutlookEX.au3>
#include "H:\tools\AutoIt3\OutlookEX\OutlookEX.au3"

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Select a sent meeting request and display the response options plus count.
; *****************************************************************************
If MsgBox($MB_OKCANCEL, "OutlookEX UDF: _OL_MeetingResponseResults Example Script", "Please select a sent meeting request in Outlook") <> $IDOK Then Exit

Global $aResponseResults = _OL_MeetingResponseResults($oOutlook, $oOutlook.ActiveExplorer.Selection(1), Default, False)
If @error Then Exit MsgBox($MB_IconError, "OutlookEX UDF: _OL_MeetingResponseResults", "Returned @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aResponseResults, "OutlookEX UDF: _OL_MeetingResponseResults - voting option and count")

; *****************************************************************************
; Example 2
; Display the verbose results - recipient and selected voting option.
; *****************************************************************************
$aResponseResults = _OL_MeetingResponseResults($oOutlook, $oOutlook.ActiveExplorer.Selection(1), Default, True)
If @error Then Exit MsgBox($MB_IconError, "OutlookEX UDF: _OL_MeetingResponseResults", "Returned @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aResponseResults, "OutlookEX UDF: _OL_MeetingResponseResults - voting option and count")
