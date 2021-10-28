#include <OutlookEX.au3>
#include <MsgBoxConstants.au3>

; *****************************************************************************
; Connect to Outlook
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF - _OL_RecipientCheck Example Script", "Error running _OL_Open. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Test recipients: The current user, a SMTP mail address and a wrong address.
; *****************************************************************************
Global $aResult = _OL_ItemRecipientCheck($oOutlook, $oOutlook.GetNameSpace("MAPI").CurrentUser.Name & ";test.user@google.com;Wrong address", "", "", "", "", "", "", "", "", "", False)
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemRecipientCheck Example Script", "Error running _OL_ItemRecipientCheck. @error = " & @error & ", @extended = " & @extended)
_ArrayDisplay($aResult, "OutlookEX UDF: _OL_RecipientCheck Example Script", _
	"", 0, "|", "Recipient|Resolved?|Recipient object|AddressEntry object|Mail address|DisplayType|Name")

; *****************************************************************************
; Example 2
; Test recipients: The current user, a SMTP mail address and a wrong address.
; Set flag $bOnlyValid = True so only valid recipients will be returned.
; @extended holds the number of invalid recipients.
; *****************************************************************************
$aResult = _OL_ItemRecipientCheck($oOutlook, $oOutlook.GetNameSpace("MAPI").CurrentUser.Name & ";test.user@google.com;Wrong address", "", "", "", "", "", "", "", "", "", True)
Global $iUnresolved = @extended
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemRecipientCheck Example Script", "Error running _OL_ItemRecipientCheck. @error = " & @error & ", @extended = " & @extended)
MsgBox($MB_ICONINFORMATION, "OutlookEX UDF: _OL_ItemRecipientCheck Example Script", "Resolved recipients: " & UBound($aResult, 1) & @CRLF & "Unresolved recipients: " & $iUnresolved)

_OL_Close($oOutlook)