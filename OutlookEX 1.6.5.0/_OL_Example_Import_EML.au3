#AutoIt3Wrapper_Run_Au3Check=Y
#Au3Stripper_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#AutoIt3Wrapper_Au3stripper_OnError=S

; *****************************************************************************
; Example Script
; Import EML-files to your Inbox.
; *****************************************************************************
#include <OutlookEX.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>

; ========== Please change the following variables to suit your environment ==========
Global $bLogging = True                                           ; Logging of information and error messages. Possible values: True, False
Global $sLoggingPath = @ScriptDir & "\_OL_Example_Import_EML.log" ; Path and filename of the log e.g. Default: @ScriptDir & "\_OL_Example_Import_EML.log"
Global $bClearLog = True                                          ; Delete an existing log file so a fresh log file gets written. Possible values: True, False
Global $bShowLog = True                                           ; Show log at end of processing. Possible values: True, False
Global $bOnErrorContinue = True                                   ; Continues processing the next EML-file when an error occurs. Possible values: True, False
Global $sDirectory = "C:\Local\EML-Test"                          ; Directory from where to import the EML-files from. If empty you will be asked
Global $iSleep = 0                                                ; In case of problems you might need to set this value >= 1000 (1 second). Outlook needs some time to process an EML-file.
Global $bSplash = True                                            ; Display a SplashText to show the progress of the script. Possible values: True, False
Global $sFolder = "*\Posteingang\test"                            ; Folder where to store the imported EML-files. Folderpath as expected by function _OL_FolderAccess. If empty you will be asked
; ========== No changes after this line please ==========

; ===================================================================
; Description:
; Script to import EML-files.
;
; Comment:
; Before executing the Script, make sure that Outlook is configured to open EML-files.
; Depending on the performance of your computer, you may need to increase the Sleep value to give Outlook
; more time to open the EML-file.
;
; author  : Robert Sparnaaij
; modified: water
; version : 1.0
; website : http://www.howto-outlook.com/howto/import-eml-files.htm
; ===================================================================
OnAutoItExitRegister("_Exit")
AutoItSetOption("TrayIconDebug", 1)

Global $sTitle = "Example script to import EML-files", $oOL, $oItem, $oFolder, $aFolder, $sDirectory, $aEMLFiles, $hTimer, $sOnError, $hSplash
Global $aFlags[][] = [["I", $MB_ICONINFORMATION], ["W", $MB_ICONWARNING], ["E", $MB_ICONERROR]]

If $bSplash Then $hSplash = SplashTextOn($sTitle, "Preparing ... please be patient!", 800, 45, 0, 0, $DLG_TEXTLEFT)

If $bLogging Then
	If $sLoggingPath = "" Then $sLoggingPath = @ScriptDir & "\_OL_Example_Import_EML.log"
	If $bClearLog Then FileDelete($sLoggingPath)
	_WriteLog(1, 0, "$bLogging........: " & $bLogging)
	_WriteLog(1, 0, "$sLoggingPath....: " & $sLoggingPath)
	_WriteLog(1, 0, "$bClearLog.......: " & $bClearLog)
	_WriteLog(1, 0, "$bShowLog........: " & $bShowLog)
	_WriteLog(1, 0, "$bOnErrorContinue: " & $bOnErrorContinue)
EndIf

If $bSplash Then ControlSetText($hSplash, "", "Static1", "Selecting import directory.")
If $sDirectory = "" Then
	$sDirectory = FileSelectFolder("Please select the directory containing the EML-files to import.", "", 0, @ScriptDir)
	If @error Then
		If @error = 1 Then ; FileSelectFolder cancelled by user
			_WriteLog(0, 1, "User cancelled the selection of the import directory. Exiting.")
		Else
			_WriteLog(0, 2, "FileSelectFolder returned an error. @error = " & @error & ". Exiting")
		EndIf
		Exit
	EndIf
EndIf
_WriteLog(1, 0, "$sDirectory......: " & $sDirectory)
_WriteLog(1, 0, "$iSleep..........: " & $iSleep)
_WriteLog(1, 0, "$bSplash.........: " & $bSplash)

If $bSplash Then ControlSetText($hSplash, "", "Static1", "Connecting to Outlook.")
$oOL = _OL_Open()
If @error Then Exit _WriteLog(0, 2, "_OL_Open returned an error. @error = " & @error & ". Exiting.")

If $bSplash Then ControlSetText($hSplash, "", "Static1", "Selecting Outlook folder to store imported EML-files.")
If $sFolder = "" Then
	$oFolder = $oOL.Session.PickFolder
	If @error Or IsObj($oFolder) = 0 Then Exit _WriteLog(0, 2, "Outlook PickFolder returned an error. @error = " & @error & ". Exiting.")
	_WriteLog(1, 0, "Picked folder....: " & $oFolder.FolderPath)
Else
	$aFolder = _OL_FolderAccess($oOL, $sFolder)
	If @error Then Exit _WriteLog(0, 2, "_OL_FolderAccess returned an error. @error = " & @error & "@extended = " & @extended & ". Exiting.")
	$oFolder = $aFolder[1]
	_WriteLog(1, 0, "$sFolder.........: " & $sFolder)
EndIf

If $bSplash Then ControlSetText($hSplash, "", "Static1", "Retrieving list of EML-files to import.")
$aEMLFiles = _FileListToArray($sDirectory, "*.eml", $FLTA_FILES, True)
If @error Then
	If @error = 4 Then
		_WriteLog(0, 1, "No EML-files found in " & $sDirectory & ". Exiting.")
	Else
		_WriteLog(0, 2, "FileListToArray returned an error. @error = " & @error & ". Exiting.")
	EndIf
	Exit
EndIf
_WriteLog(1, 0, "Number of EML-files to process: " & $aEMLFiles[0])

$sOnError = ($bOnErrorContinue = True) ? ". Continuing." : ". Exiting."
If $bSplash Then ControlSetText($hSplash, "", "Static1", "Start importing.")
$hTimer = TimerInit()
For $i = 1 To $aEMLFiles[0]
	If $bSplash Then ControlSetText($hSplash, "", "Static1", "Importing EML-file " & $i & " of " & $aEMLFiles[0] & ": " & $aEMLFiles[$i])
	ShellExecuteWait($aEMLFiles[$i], "", "", "open", 1)
	If @error Then
		_WriteLog(0, 2, "ShellExecuteWait returned an error when importing EML-file " & $aEMLFiles[$i] & ". @error = " & @error & $sOnError)
		If $bOnErrorContinue Then ContinueLoop
		Exit
	EndIf
	Sleep($iSleep)
	$oItem = $oOL.ActiveInspector.CurrentItem
	If @error Then
		_WriteLog(0, 2, "Could not get the Outlook Inspector for EML-file " & $aEMLFiles[$i] & ". @error = " & @error & $sOnError)
		If $bOnErrorContinue Then ContinueLoop
		Exit
	EndIf
	$oItem.Move($oFolder)
	If @error Then
		_WriteLog(0, 2, "Could not move the created mail item for EML-file " & $aEMLFiles[$i] & ". @error = " & @error & $sOnError)
		If $bOnErrorContinue Then ContinueLoop
		Exit
	EndIf
	_WriteLog(1, 0, "File " & $i & " of " & $aEMLFiles[0] & " processed: " & $aEMLFiles[$i])
Next
_WriteLog(0, 0, "Finished processing " & $aEMLFiles[0] & " EML-files after " & StringFormat("%.2f", TimerDiff($hTimer)/1000) & " seconds.")
If $bSplash Then ControlSetText($hSplash, "", "Static1", "Done.")
Exit

Func _WriteLog($iFlag, $iSeverity, $sMessage)
	Local $sSeverity = $aFlags[$iSeverity][0]
	Local $MBFlag = $aFlags[$iSeverity][1]
	If $iFlag = 1 Then
		If $bLogging Then FileWriteLine($sLoggingPath, @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " " & $sSeverity & " " & $sMessage)
	Else
		If $bLogging Then
			FileWriteLine($sLoggingPath, @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " " & $sSeverity & " " & $sMessage)
		Else
			MsgBox($MBFlag, $sTitle, $sMessage)
		EndIf
	EndIf
EndFunc   ;==>_WriteLog

Func _Exit()
	If $bLogging And $bShowLog Then Run("Notepad " & $sLoggingPath)
EndFunc   ;==>_Exit

#cs
'===================================================================
'Description: VBS script to import eml-files.
'
'Comment: Before executing the vbs-file, make sure that Outlook is
'         configured to open eml-files.
'         Depending on the performance of your computer, you may
'         need to increase the Wscript.Sleep value to give Outlook
'         more time to open the eml-file.
'
' author : Robert Sparnaaij
' version: 1.0
' website: http://www.howto-outlook.com/howto/import-eml-files.htm
'===================================================================

Dim objShell : Set objShell = CreateObject("Shell.Application")
Dim objFolder : Set objFolder = objShell.BrowseForFolder(0, "Select the folder containing eml-files", 0)

Dim Item
If (NOT objFolder is Nothing) Then
  Set WShell = CreateObject("WScript.Shell")
  Set objOutlook = CreateObject("Outlook.Application")
  Set Folder = objOutlook.Session.PickFolder
  If NOT Folder Is Nothing Then
    For Each Item in objFolder.Items
      If Right(Item.Name, 4) = ".eml" AND Item.IsFolder = False Then
	objShell.ShellExecute Item.Path, "", "", "open", 1
	WScript.Sleep 1000
	Set MyInspector = objOutlook.ActiveInspector
	Set MyItem = objOutlook.ActiveInspector.CurrentItem
	MyItem.Move Folder
      End If
    Next
  End If
End If

MsgBox "Import completed.", 64, "Import EML"

Set objFolder = Nothing
Set objShell = Nothing﻿
#ce