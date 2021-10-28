#include <OutlookEX.au3>

Global $oOL = _OL_Open()
Global $aFolder = _OL_FolderAccess($oOL, Default, $olFolderInbox)
If @error Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemAccessGet Example Script", "Error accessing the Inbox. @error = " & @error & ", @extended = " & @extended)
Sleep(4000)
; *****************************************************************************
; Example 1
; Get the access rights for the users Inbox
; *****************************************************************************
$iAccess = _OL_ItemAccessGet($aFolder[1], 1)
If @error Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemAccessGet Example Script", "Error retrieving the ACCESS_LEVEL property of the Inbox. @error = " & @error & ", @extended = " & @extended)
MsgBox($MB_ICONINFORMATION, "_OL_ItemAccessGet Example Script", "ACCESS LEVEL Property of the Inbox: " & $iAccess)

$iAccess = _OL_ItemAccessGet($aFolder[1], 2)
If @error Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemAccessGet Example Script", "Error retrieving the ACCESS property of the Inbox. @error = " & @error & ", @extended = " & @extended)
Local $sMessage = ""
If BitAND($iAccess, 0x00000001) = 0x00000001 Then $sMessage &= "  0x1: Write" & @CRLF
If BitAND($iAccess, 0x00000002) = 0x00000002 Then $sMessage &= "  0x2: Read" & @CRLF
If BitAND($iAccess, 0x00000004) = 0x00000004 Then $sMessage &= "  0x4: Delete" & @CRLF
If BitAND($iAccess, 0x00000008) = 0x00000008 Then $sMessage &= "  0x8: Create subfolders in the folder hierarchy" & @CRLF
If BitAND($iAccess, 0x00000010) = 0x00000010 Then $sMessage &= "  0x10: Create content messages" & @CRLF
If BitAND($iAccess, 0x00000020) = 0x00000020 Then $sMessage &= "  0x20: Create associated content messages" & @CRLF
MsgBox($MB_ICONINFORMATION, "_OL_ItemAccessGet Example Script", "ACCESS Property of the Inbox: " & @CRLF & $sMessage)

; *****************************************************************************
; Example 2
; Get the access rights of the first item in the users Inbox
; *****************************************************************************
$iAccess = _OL_ItemAccessGet($aFolder[1].Items.Item(1), 1)
If @error Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemAccessGet Example Script", "Error retrieving the ACCESS_LEVEL property of the first item in the Inbox. @error = " & @error & ", @extended = " & @extended)
MsgBox($MB_ICONINFORMATION, "_OL_ItemAccessGet Example Script", "ACCESS LEVEL Property of the first item in the Inbox: " & $iAccess)

$iAccess = _OL_ItemAccessGet($aFolder[1].Items.Item(1), 2)
If @error Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemAccessGet Example Script", "Error retrieving the ACCESS property of the first item in the Inbox. @error = " & @error & ", @extended = " & @extended)
Local $sMessage = ""
If BitAND($iAccess, 0x00000001) = 0x00000001 Then $sMessage &= "  0x1: Write" & @CRLF
If BitAND($iAccess, 0x00000002) = 0x00000002 Then $sMessage &= "  0x2: Read" & @CRLF
If BitAND($iAccess, 0x00000004) = 0x00000004 Then $sMessage &= "  0x4: Delete" & @CRLF
If BitAND($iAccess, 0x00000008) = 0x00000008 Then $sMessage &= "  0x8: Create subfolders in the folder hierarchy" & @CRLF
If BitAND($iAccess, 0x00000010) = 0x00000010 Then $sMessage &= "  0x10: Create content messages" & @CRLF
If BitAND($iAccess, 0x00000020) = 0x00000020 Then $sMessage &= "  0x20: Create associated content messages" & @CRLF
MsgBox($MB_ICONINFORMATION, "_OL_ItemAccessGet Example Script", "ACCESS Property of the first item in the Inbox: " & @CRLF & $sMessage)

_OL_Close($oOL)