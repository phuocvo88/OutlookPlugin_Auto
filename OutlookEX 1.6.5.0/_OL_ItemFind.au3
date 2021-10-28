#include <OutlookEX.au3>

; *****************************************************************************
; Create test environment
; *****************************************************************************
Global $oOutlook = _OL_Open()
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF", "Error creating a connection to Outlook. @error = " & @error & ", @extended = " & @extended)
Global $Result = _OL_TestEnvironmentCreate($oOutlook)
If @error <> 0 Then Exit MsgBox($MB_ICONERROR, "OutlookEX UDF - Manage Test Environment", "Error creating the test environment. @error = " & @error & ", @extended = " & @extended)
Global $aItems, $iItems

;------------------------------------------------------------------------------------------------------------------------------------------------
; Example 1 - Search for contacts with firstname = TestFirstName in the contacts folder and every sub-folder returning default properties
;------------------------------------------------------------------------------------------------------------------------------------------------
$aItems = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", $olContact, '[FirstName] = "TestFirstName"', "", "", "", "", 1)
If @error = 0 Then
	_ArrayDisplay($aItems, "OutlookEX UDF: _OL_ItemFind Example Script - Find contacts by firstname")
Else
	MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemFind Example Script", "Error finding a contact. @error = " & @error & ", @extended: " & @extended)
EndIf

;------------------------------------------------------------------------------------------------------------------------------------------------
; Example 2 - Search for appointments with "Room" as location (partial match) in the calendar folder and every subfolder
;------------------------------------------------------------------------------------------------------------------------------------------------
$aItems = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Calendar", $olAppointment, "", "Location", "Room", "EntryID,Subject,Location", "", 1)
If @error = 0 Then
	_ArrayDisplay($aItems, "OutlookEX UDF: _OL_ItemFind Example Script - Find appointments by partial search")
Else
	MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemFind Example Script", "Error finding an appointment. @error = " & @error & ", @extended: " & @extended)
EndIf

;------------------------------------------------------------------------------------------------------------------------------------------------
; Example 3 - Get number of items (contacts without distribution lists) in the contacts folder
;------------------------------------------------------------------------------------------------------------------------------------------------
$iItems = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Contacts", $olContact, "", "", "", "", "", 4)
If @error = 0 Then
	MsgBox($MB_ICONINFORMATION, "OutlookEX UDF: _OL_ItemFind Example Script", "Number of items found: " & $iItems)
Else
	MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemFind Example Script", "Could not find an item in the contacts folders. @error = " & @error & ", @extended: " & @extended)
EndIf

;------------------------------------------------------------------------------------------------------------------------------------------------
; Example 4 - Get unread mails from a folder and all subfolders
;             Returning Subject, Body and two pseudo properties (object of the folder and item)
;------------------------------------------------------------------------------------------------------------------------------------------------
$aItems = _OL_ItemFind($oOutlook, "*\Outlook-UDF-Test\SourceFolder\Mail", $olMail, "[UnRead]=True", "", "", "Subject,Body,@FolderObject,@ItemObject", "", 2)
If @error = 0 Then
	_ArrayDisplay($aItems, "OutlookEX UDF: _OL_ItemFind Example Script - Unread mails")
	MsgBox($MB_ICONINFORMATION, "OutlookEX UDF: _OL_ItemFind Example Script", "Full path of the folder where the first item was found: " & @CRLF & $aItems[1][2].FolderPath)
Else
	MsgBox($MB_ICONERROR, "OutlookEX UDF: _OL_ItemFind Example Script", "Could not find an unread mail. @error = " & @error & ", @extended: " & @extended)
EndIf

_OL_Close($oOutlook)