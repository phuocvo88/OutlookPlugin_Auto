#include <OutlookEX.au3>
#include <MsgBoxConstants.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)
Global $Result = _OL_TestEnvironmentCreate($oOutlook)
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error & ", @extended = " & @extended)
Global $sCurrentUser = $oOutlook.GetNameSpace("MAPI").CurrentUser.Name

; *****************************************************************************
; Example 1
; Add an optional recipient (the current user) to a meeting
; *****************************************************************************
Global $aOL_Item = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Calendar", $olAppointment, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemRecipientAdd Example Script", "Could not find an appointment item in folder 'Outlook-UDF-Test\SourceFolder\Calendar'. @error = " & @error)
Global $oItem = _OL_ItemRecipientAdd($oOutlook, $aOL_Item[1][0], Default, $olOptional, $sCurrentUser)
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemRecipientAdd Example Script", "Error adding recipient to appointment in folder 'Outlook-UDF-Test\SourceFolder\Calendar'. @error = " & @error & ", @extended = " & @extended)

; Display item
_OL_ItemDisplay($oOutlook, $oItem)
MsgBox($MB_ICONINFORMATION, "OutlookEX UDF: _OL_ItemRecipientAdd Example Script", "Recipient successfully added to the appointment!")

; *****************************************************************************
; Example 2
; Add a recipient and two reply-recipients to a mail item
; *****************************************************************************
; Create HTML mail item
$oItem = _OL_ItemCreate($oOutlook, $olMailItem, "*\Outlook-UDF-Test\TargetFolder\Mail", "", "Subject=TestMail", "BodyFormat=" & $olFormatHTML)
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemRecipientAdd Example Script", "Error creating a mail item. @error = " & @error & ", @extended = " & @extended)

; Add a recipient
_OL_ItemRecipientAdd($oOutlook, $oItem, Default, $olTo, $sCurrentUser)
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemRecipientAdd Example Script", "Error adding recipient to mail item in folder 'Outlook-UDF-Test\TargetFolder\Mail'. @error = " & @error & ", @extended = " & @extended)

; Add two reply-recipients
_OL_ItemRecipientAdd($oOutlook, $oItem, Default, $olReplyRecipient, "John.Doe@gmx.com")
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemRecipientAdd Example Script", "Error adding reply-recipient to mail item in folder 'Outlook-UDF-Test\TargetFolder\Mail'. @error = " & @error & ", @extended = " & @extended)

_OL_ItemRecipientAdd($oOutlook, $oItem, Default, $olReplyRecipient, "Jane.Doe@gmx.com")
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemRecipientAdd Example Script", "Error adding reply-recipient to mail item in folder 'Outlook-UDF-Test\TargetFolder\Mail'. @error = " & @error & ", @extended = " & @extended)

; Display item
_OL_ItemDisplay($oOutlook, $oItem)
MsgBox($MB_ICONINFORMATION, "OutlookEX UDF: _OL_ItemRecipientAdd Example Script", "Recipient and reply-recipients successfully added to the mail item!")

_OL_Close($oOutlook)