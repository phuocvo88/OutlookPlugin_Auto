#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oItem
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)
Global $Result = _OL_TestEnvironmentCreate($oOutlook)
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Create a mail with voting buttons (but don't send it)
; *****************************************************************************
$oItem = _OL_ItemCreate($oOutlook, $olMailItem, "*\Outlook-UDF-Test\TargetFolder\Mail", "", "Subject=TestMail", "Body=Testvoting!")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_MailVotingSet Example Script", "Error creating a mail in folder 'Outlook-UDF-Test\TargetFolder\Mail'. @error = " & @error & ", @extended = " & @extended)

_OL_MailVotingSet($oItem, "Red|Blue|Green|White", "Green")
$oItem.Display
MsgBox(64, "OutlookEX UDF: _OL_MailVotingSet Example Script", "Mail with voting buttons created.")

_OL_Close($oOutlook)