#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Select a sent mail with a voting and display the voting options plus count.
; In addition use a different text for the "No Reply" option.
; *****************************************************************************
If MsgBox($MB_OKCANCEL, "OutlookEX UDF: _OL_MailVotingResults Example Script", "Please select a sent mail item with a voting in Outlook") <> $IDOK Then Exit

Global $aVotingResults = _OL_MailVotingResults($oOutlook, $oOutlook.ActiveExplorer.Selection(1), Default, False, "Didn't get a reply from this guys")
If @error Then Exit MsgBox($MB_IconError, "OutlookEX UDF: _OL_MailVotingResults", "Returned @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aVotingResults, "OutlookEX UDF: _OL_MailVotingResults - voting option and count")

; *****************************************************************************
; Example 2
; Display the verbose results - recipient and selected voting option.
; *****************************************************************************
$aVotingResults = _OL_MailVotingResults($oOutlook, $oOutlook.ActiveExplorer.Selection(1), Default, True)
If @error Then Exit MsgBox($MB_IconError, "OutlookEX UDF: _OL_MailVotingResults", "Returned @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aVotingResults, "OutlookEX UDF: _OL_MailVotingResults - voting option and count")
