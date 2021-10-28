#include <OutlookEX_GUI.au3>
#include <File.au3>

; *****************************************************************************
; Example Script
; Restores previously backed up favorites to the mail navigation folder.
; *****************************************************************************
Global $aGroups[0][0], $iResult, $oOL
Global $sTitle = "Outlook Favorites - Restore"
Global $sBackupFile = @ScriptDir & "\Outlook Favorites - Backup.txt"

; *****************************************************************************
; Connect to Outlook
; *****************************************************************************
$oOL = _OL_Open()
If @error <> 0 Then Exit MsgBox($MB_IconError, $sTitle, "Error connecting to Outlook. @error = " & @error & ", @extended = " & @extended)

; Translate folderpath to the format used by _OL_FolderAccess
_FileReadToArray($sBackupFile, $aGroups, $FRTA_NOCOUNT, "|")
For $i = 0 To UBound($aGroups) - 1
	$aGroups[$i][2] = StringMid($aGroups[$i][2], 3)
Next

; *****************************************************************************
; Add the favorites from the backup file
; *****************************************************************************
$iResult = _OL_NavigationFolderAdd($oOL, $aGroups)
If @error <> 0 Then Exit MsgBox($MB_IconError, $sTitle, "Error adding favorites to the navigation folder. @error = " & @error & ", @extended = " & @extended)
MsgBox($MB_IconInformation, $sTitle, UBound($aGroups) & " items restored from " & @CRLF & $sBackupFile & ".", 10)

_OL_Close($oOL)