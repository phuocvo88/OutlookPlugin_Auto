#include <OutlookEX_GUI.au3>
#include <MsgBoxConstants.au3>
#include <File.au3>

; *****************************************************************************
; Example Script
; Backup your favorites from the mail navigation folder.
; *****************************************************************************
Global $aGroups[0][0], $oOL
Global $sTitle = "Outlook favorites - Backup"
Global $sBackupFile = @ScriptDir & "\Outlook Favorites - Backup.txt"

; *****************************************************************************
; Connect to Outlook
; *****************************************************************************
$oOL = _OL_Open()
If @error <> 0 Then Exit MsgBox($MB_IconError, $sTitle, "Error connecting to Outlook. @error = " & @error & ", @extended = " & @extended)

; *****************************************************************************
; Get a list of all navigation groups and folders in the mail navigation module
; *****************************************************************************
$aGroups = _OL_NavigationFolderGet($oOL, $olModuleMail)
If @error <> 0 Then Exit MsgBox($MB_IconError, $sTitle, "Error getting navigation groups / folders of the mail navigation module. @error = " & @error & ", @extended = " & @extended)
; Write the results to a text file
_FileWriteFromArray($sBackupFile, $aGroups, 1)
MsgBox($MB_IconInformation, $sTitle, $aGroups[0][0] & " items saved to " & @CRLF & $sBackupFile & ".", 10)

_OL_Close($oOL)