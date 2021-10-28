#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)

; Global $Result = _OL_TestEnvironmentCreate($oOutlook)
; If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Example 1
; Find a mail item in the specified folder and create 20 copies in the same folder
; Measure the time needed when using _OL_ItemCopy and _OL_ItemBulk
; _OL_ItemBulk returns errors as
; *****************************************************************************
Global $aFolder = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail")
Global $aOL_Item = _OL_ItemFind($oOutlook, $aFolder[1], $olMail, "", "", "", "EntryID")
If $aOL_Item[0][0] = 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemBulk Example Script", "Could not find a mail item in folder '*\Outlook-UDF-Test\SourceFolder\Mail'. @error = " & @error)
ReDim $aOL_Item[22][1]
; #cs
; Use the item object instead of the EntryID
$aOL_Item[1][0] = _OL_ItemGet($oOutlook, $aOL_Item[1][0], Default, -1)
For $i = 1 to 20
	$aOL_Item[$i+1][0] = $aOL_Item[$i][0]
Next

Global $iTime = TimerInit()
for $i = 1 to 20
	_OL_ItemCopy($oOutlook, $aOL_Item[$i][0])
Next
MsgBox(64, "OutlookEX UDF: _OL_ItemBulk Example Script", "Mails successfully copied using _OL_ItemCopy! It took " & TimerDiff($iTime) & " milliseconds.")
; #ce

; #cs
; Copy
$aFolder = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail\Ziel")
Global $iFlag = 1
$iTime = TimerInit()
Global $aErrors = _OL_ItemBulk($oOutlook, $aOL_Item, Default, $aFolder[1], 1, Default, Default, $iFlag)
If @error <> 0 Then
	MsgBox(16, "OutlookEX UDF: _OL_ItemBulk Example Script", "Error copying specified items. @error = " & @error & ", @extended = " & @extended)
	If $iFlag = 1 Then _ArrayDisplay($aErrors)
Else
	MsgBox(64, "OutlookEX UDF: _OL_ItemBulk Example Script", "Mails successfully copied using _OL_ItemBulk ! It took " & TimerDiff($iTime) & " milliseconds.")
EndIf
; #ce

#cs
; Move
$aFolder = _OL_FolderAccess($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail\Ziel")
Global $iFlag = 1
$iTime = TimerInit()
Global $aErrors = _OL_ItemBulk($oOutlook, $aOL_Item, Default, $aFolder[1], 3, Default, Default, $iFlag)
If @error <> 0 Then
	MsgBox(16, "OutlookEX UDF: _OL_ItemBulk Example Script", "Error moving specified items. @error = " & @error & ", @extended = " & @extended)
	If $iFlag = 1 Then _ArrayDisplay($aErrors)
Else
	MsgBox(64, "OutlookEX UDF: _OL_ItemBulk Example Script", "Mails successfully moved using _OL_ItemBulk ! It took " & TimerDiff($iTime) & " milliseconds.")
EndIf
#ce

_OL_Close($oOutlook)