#include <OutlookEX.au3>

; *****************************************************************************
; Connect to Outlook
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_MailheaderGet Example Script", "Error connecting to Outlook. @error = " & @error & ", @extended: " & @extended)

; *****************************************************************************
; Example 1
; Get the first unread mail in the inbox and display the mail headers
; *****************************************************************************
; Access the inbox
Global $aFolder = _OL_FolderAccess($oOutlook, "", $olFolderInbox)
If @error Then Exit MsgBox(48, "", "@error = " & @error & ", @extended: " & @extended)
; Find all unread items in the inbox
Global $aItems = _OL_ItemFind($oOutlook, $aFolder[1], $olMail, "[UnRead]=True", "", "", "EntryID,Subject", "", 0)
If @error Or Not IsArray($aItems) Then Exit MsgBox(48, "OutlookEX UDF: _OL_MailheaderGet Example Script", "_OL_ItemFind: @error = " & @error & ", @extended: " & @extended)
; Get the mail headers of the first mail
Global $sMailHeaders = _OL_MailheaderGet($oOutlook, $aItems[1][0])
If @error Then Exit MsgBox(16, "OutlookEX UDF: _OL_MailheaderGet Example Script", "Error retrieving mail headers of mail with subject '" & $aItems[1][1] & "'. @error = " & @error & ", @extended: " & @extended)
MsgBox(64, "OutlookEX UDF: _OL_MailheaderGet Example Script", "Mail headers of mail with subject '" & $aItems[1][1] & "'." & @CRLF & @CRLF & $sMailHeaders)

_OL_Close($oOutlook)